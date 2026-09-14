# utils/image_annotator.py
"""
交互式图像标注工具
用于在功能架构图上标注优化模块
"""

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QListWidget, QListWidgetItem, QLineEdit, QMessageBox,
    QScrollArea, QWidget, QTextEdit, QFrame, QComboBox, QCheckBox,
    QGroupBox, QFileDialog
)
from PySide6.QtCore import Qt, QRect, QPoint, Signal
from PySide6.QtGui import QPixmap, QPainter, QPen, QColor, QFont, QImage
import os
import json
from utils.runtime_logger import log_error, log_info, log_warn


class OCRConfig:
    """OCR配置管理"""
    CONFIG_FILE = "ocr_config.json"
    
    @staticmethod
    def load():
        """加载配置"""
        default_config = {
            "enabled_engines": ["paddleocr", "easyocr", "tesseract"],
            "tesseract_path": "",
            "enable_image_enhance": True
        }
        
        try:
            if os.path.exists(OCRConfig.CONFIG_FILE):
                with open(OCRConfig.CONFIG_FILE, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # 如果配置文件里的引擎列表为空，则强制使用默认全选
                    if not config.get("enabled_engines"):
                        config["enabled_engines"] = ["paddleocr", "easyocr", "tesseract"]
                    # 合并默认配置
                    default_config.update(config)
        except Exception as e:
            log_error(f'[OCR] 加载配置失败: {e}，使用默认配置')
        
        return default_config
    
    @staticmethod
    def save(config):
        """保存配置"""
        try:
            with open(OCRConfig.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            log_info(f'[OCR] 配置已保存到 {OCRConfig.CONFIG_FILE}')
        except Exception as e:
            log_error(f'[OCR] 保存配置失败: {e}')


class AnnotationLabel(QLabel):
    """可交互的图像标签，支持点击和框选"""
    
    annotation_added = Signal(QRect, str, str)  # 区域, 模块编号, 模块名称
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.original_pixmap = None
        self.annotations = []  # [(rect, num, name, color), ...]
        self.current_rect = None
        self.start_point = None
        self.is_drawing = False
        self.scale_factor = 1.0  # 缩放因子
        self.min_scale = 0.1
        self.max_scale = 5.0
        self.offset = QPoint(0, 0)
        
        # 拖动相关状态
        self.is_dragging = False
        self.drag_start_pos = None
        self.drag_start_offset = QPoint(0, 0)
        
        # OCR配置
        self.enabled_engines = ["paddleocr", "easyocr", "tesseract"]  # 默认优先 PaddleOCR
        self.tesseract_path = "D:/IT/Tesseract-OCR/tesseract.exe"
        self.enable_image_enhance = True
        
        self.setMouseTracking(True)
        self.setAlignment(Qt.AlignCenter)
        
    def load_image(self, image_path):
        """加载图片"""
        self.original_pixmap = QPixmap(image_path)
        if self.original_pixmap.isNull():
            return False
        
        # 计算初始缩放比例，适应窗口
        # 但保留用户手动缩放的能力
        max_width = 1000
        max_height = 700
        width_scale = max_width / self.original_pixmap.width()
        height_scale = max_height / self.original_pixmap.height()
        self.scale_factor = min(width_scale, height_scale, 1.0)  # 不放大，只缩小
        
        self.update_display()
        return True
    
    def mousePressEvent(self, event):
        """鼠标按下 - 左键框选，右键/中键拖动"""
        if not self.original_pixmap:
            return
        
        # 彻底重置状态
        self.is_drawing = False
        self.is_dragging = False
        
        if event.button() == Qt.LeftButton:
            pos = self.map_to_image(event.pos())
            if pos:
                self.start_point = pos
                self.is_drawing = True
                self.current_rect = None
                self.setCursor(Qt.CrossCursor)
        
        elif event.button() in (Qt.RightButton, Qt.MiddleButton):
            self.is_dragging = True
            self.drag_start_pos = event.pos()
            self.drag_start_offset = self.offset
            self.setCursor(Qt.ClosedHandCursor)
        
        self.update()

    def mouseMoveEvent(self, event):
        """鼠标移动 - 更新框选区域或拖动图像"""
        if self.is_drawing and self.start_point:
            # 框选模式
            pos = self.map_to_image(event.pos())
            if pos:
                self.current_rect = QRect(self.start_point, pos).normalized()
                # 重要：直接调用 update 触发 paintEvent，避免 setPixmap 导致的布局抖动
                self.update()
            return
        
        if self.is_dragging and self.drag_start_pos:
            # 拖动模式
            delta = event.pos() - self.drag_start_pos
            self.offset = self.drag_start_offset + delta
            self.update()
            return

    def mouseReleaseEvent(self, event):
        """鼠标释放 - 完成框选或拖动"""
        if event.button() == Qt.LeftButton and self.is_drawing:
            self.is_drawing = False
            if self.current_rect and self.current_rect.width() > 10 and self.current_rect.height() > 10:
                self.show_module_input_dialog(self.current_rect)
            self.current_rect = None
            self.start_point = None
        
        elif event.button() in (Qt.RightButton, Qt.MiddleButton):
            self.is_dragging = False
            self.drag_start_pos = None
            
        self.setCursor(Qt.CrossCursor)
        self.update()
    
    def wheelEvent(self, event):
        """鼠标滚轮事件 - 缩放图片"""
        if not self.original_pixmap:
            return
        
        # 获取滚轮角度
        delta = event.angleDelta().y()
        
        # 计算缩放因子变化
        if delta > 0:
            # 向上滚动 - 放大
            scale_change = 1.1
        else:
            # 向下滚动 - 缩小
            scale_change = 0.9
        
        new_scale = self.scale_factor * scale_change
        
        # 限制缩放范围
        if self.min_scale <= new_scale <= self.max_scale:
            self.scale_factor = new_scale
            self.update_display()
            
            # 显示缩放比例提示
            if hasattr(self.parent(), 'show_scale_tip'):
                self.parent().show_scale_tip(int(self.scale_factor * 100))
    
    def map_to_image(self, pos):
        """将屏幕坐标转换为图像坐标（考虑缩放和拖动偏移）"""
        if not self.original_pixmap:
            return None
        
        # 获取当前显示的图像大小
        scaled_width = int(self.original_pixmap.width() * self.scale_factor)
        scaled_height = int(self.original_pixmap.height() * self.scale_factor)
        
        widget_rect = self.rect()
        
        # 计算图像在控件中的位置（居中 + offset）
        x_offset = (widget_rect.width() - scaled_width) // 2 + self.offset.x()
        y_offset = (widget_rect.height() - scaled_height) // 2 + self.offset.y()
        
        # 转换坐标
        x = (pos.x() - x_offset) / self.scale_factor
        y = (pos.y() - y_offset) / self.scale_factor
        
        # 检查是否在图像范围内
        if 0 <= x < self.original_pixmap.width() and 0 <= y < self.original_pixmap.height():
            return QPoint(int(x), int(y))
        return None
    
    def show_module_input_dialog(self, rect, existing_data=None):
        """显示模块输入对话框"""
        # 如果是编辑模式，使用现有数据
        if existing_data:
            num, name, color = existing_data
            ocr_text = name
        else:
            # 尝试OCR识别框选区域的文字
            ocr_text = self.ocr_recognize_region(rect)
        
        dialog = ModuleInputDialog(self, ocr_text)
        if existing_data:
            # 设置对话框初始值
            # 注意：ModuleInputDialog 内部根据层级生成编号，所以我们这里反推层级
            level = num.count('.') + 1
            dialog.level_combo.setCurrentIndex(level - 1)
            dialog.name_input.setText(name)
            # 设置颜色（如果匹配）
            for i in range(dialog.color_combo.count()):
                if dialog.color_combo.itemData(i) == color:
                    dialog.color_combo.setCurrentIndex(i)
                    break

        if dialog.exec() == QDialog.Accepted:
            module_num, module_name, color = dialog.get_values()
            if module_name:  # 只需要名称即可
                if existing_data:
                    # 更新现有标注
                    for i, ann in enumerate(self.annotations):
                        if ann[0] == rect:
                            self.annotations[i] = (rect, module_num, module_name, color)
                            # 触发重新渲染
                            self.update_display()
                            # 发送更新信号 (可以使用相同的信号，父组件处理)
                            self.annotation_added.emit(rect, module_num, module_name)
                            break
                else:
                    self.annotations.append((rect, module_num, module_name, color))
                    self.annotation_added.emit(rect, module_num, module_name)
                    self.update_display()

    def mouseDoubleClickEvent(self, event):
        """双击编辑现有标注"""
        pos = self.map_to_image(event.pos())
        if not pos:
            return
            
        # 查找被双击的标注
        for rect, num, name, color in reversed(self.annotations):
            if rect.contains(pos):
                self.show_module_input_dialog(rect, (num, name, color))
                break
    
    def preprocess_image(self, arr, engine_type="paddleocr"):
        """图像预处理 - 为不同引擎优化的版本 (参考Umi-OCR思路)"""
        if not self.enable_image_enhance:
            return arr
        
        try:
            import cv2
            import numpy as np
            
            # 1. 转灰度
            gray = cv2.cvtColor(arr, cv2.COLOR_RGB2GRAY) if len(arr.shape) == 3 else arr
            
            # 2. 增加白边 (50px) - 确保文字不紧贴边缘，这对 OCR 引擎非常重要
            pad = 50
            gray = cv2.copyMakeBorder(gray, pad, pad, pad, pad, cv2.BORDER_CONSTANT, value=255)
            
            # 3. 智能缩放 (对于小区域大幅缩写)
            # Tesseract 最佳高度在 150-300px 左右，PaddleOCR 约 128px 最佳
            h, w = gray.shape
            target_h = 240 if engine_type == "tesseract" else 128
            
            if h < target_h:
                scale = target_h / h
                # 使用高质量 Lanczos 插值
                gray = cv2.resize(gray, (int(w * scale), int(h * scale)), interpolation=cv2.INTER_LANCZOS4)
            
            # 4. 对比度增强 (CLAHE)
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
            gray = clahe.apply(gray)
            
            # --- 核心改进：智能去边框算法 (针对架构图边框优化) ---
            # 1. 反转二值化（寻找白色背景上的黑色形状）
            _, binary = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
            
            # 2. 查找所有轮廓
            contours, _ = cv2.findContours(binary, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            
            # 3. 创建掩膜，只保留“文字样”的像素簇，剔除“边框样”的长线条
            mask = np.zeros_like(binary)
            h_img, w_img = binary.shape
            for cnt in contours:
                x_c, y_c, w_c, h_c = cv2.boundingRect(cnt)
                # 边框特征：极长、极宽或几乎顶到图片边缘
                is_border = (w_c > w_img * 0.85) or (h_c > h_img * 0.85)
                # 文字特征：高度适中，且不是那种细长横线或竖线 (虚线框通常由小段线段组成)
                if not is_border and h_c > 4 and w_c > 2:
                    cv2.drawContours(mask, [cnt], -1, 255, -1)
            
            # 4. 应用掩膜并恢复
            clean_binary = cv2.bitwise_and(binary, mask)
            # 恢复为白底黑字
            final_gray = cv2.bitwise_not(clean_binary)
            
            # --- 引擎特定微调 ---
            if engine_type == "tesseract":
                # Tesseract 喜欢高对比度
                return cv2.cvtColor(final_gray, cv2.COLOR_GRAY2RGB)
            
            elif engine_type == "paddleocr":
                # PaddleOCR 喜欢稍微柔和一点的灰度，我们在清理后的图上做轻微模糊
                final_gray = cv2.GaussianBlur(final_gray, (0, 0), 0.5)
                return cv2.cvtColor(final_gray, cv2.COLOR_GRAY2RGB)
            
            # 默认返回清理后的三通道图
            return cv2.cvtColor(final_gray, cv2.COLOR_GRAY2RGB)
            
        except Exception as e:
            log_warn(f'[OCR] ⚠️ 预处理优化失败: {e}')

            return arr
    
    def ocr_recognize_region(self, rect):
        """OCR识别指定区域的文字 - 按配置的引擎顺序尝试"""
        # 按用户启用的引擎顺序依次尝试
        for engine in self.enabled_engines:
            if engine == "easyocr":
                result = self._try_easyocr(rect)
                if result:
                    return result
            elif engine == "paddleocr":
                result = self._try_paddleocr(rect)
                if result:
                    return result
            elif engine == "tesseract":
                result = self._try_tesseract(rect)
                if result:
                    return result
        
        log_error('[OCR] 所有启用的OCR引擎都失败')
        return ""
    
    def _try_easyocr(self, rect):
        """尝试使用EasyOCR识别"""
        try:
            import easyocr
            import numpy as np
            
            log_info(f'[EasyOCR] 识别区域: {rect}')
            
            # 初始化EasyOCR（只在第一次使用时初始化）
            if not hasattr(self, '_easyocr_reader'):
                log_info('[EasyOCR] 初始化引擎（首次加载需要下载模型）...')
                self._easyocr_reader = easyocr.Reader(
                    ['ch_sim', 'en'], 
                    gpu=False,
                    verbose=False,
                    quantize=True,  # 量化加速
                    cudnn_benchmark=False
                )
                log_info('[EasyOCR] 初始化成功')
            
            # 裁剪图像
            arr = self._crop_region_to_numpy(rect)
            if arr is None:
                return None
            
            # 图像预处理 (EasyOCR 优先使用原灰度)
            arr = self.preprocess_image(arr, engine_type="easyocr")
            
            # EasyOCR识别
            result = self._easyocr_reader.readtext(
                arr, 
                detail=1,
                paragraph=False
            )
            
            if result:
                # 过滤置信度低的结果
                output_texts = []
                for res in result:
                    pos, text, conf = res
                    if conf > 0.4: # 只要不是极其模糊的就保留
                        output_texts.append(text)
                
                if output_texts:
                    text = ' '.join([s.strip() for s in output_texts if s.strip()])
                    text = self._post_process_text(text)
                    if text:
                        log_info(f'[EasyOCR] ✓ 识别结果: {text}')
                        return text
                
                log_error('[EasyOCR] ✗ 置信度过低或结果为空')
                return None
            else:
                log_error('[EasyOCR] ✗ 未识别到文字')
                return None
                
        except ImportError:
            log_error('[EasyOCR] ✗ 未安装（pip install easyocr）')
            return None
        except Exception as e:
            log_error(f'[EasyOCR] ✗ 识别失败: {e}')
            return None
    
    def _try_paddleocr(self, rect):
        """尝试使用PaddleOCR识别 (针对 Paddle 3.x 环境及 OneDNN 报错特别优化)"""
        try:
            import os
            # 1. 极端强制的环境设置：必须在 import paddle 之前或立即执行
            os.environ["FLAGS_enable_pir_api"] = "0" 
            os.environ["FLAGS_use_mkldnn"] = "0"
            os.environ["FLAGS_onednn_enabled"] = "0"
            
            import numpy as np
            from paddleocr import PaddleOCR
            
            # 2. 尝试使用 paddle 自身的 flag 设置 API (针对已加载环境)
            try:
                import paddle
                paddle.set_flags({
                    'FLAGS_use_mkldnn': 0,
                    'FLAGS_onednn_enabled': 0,
                    'FLAGS_enable_pir_api': 0
                })
            except:
                pass
            
            log_info(f'[PaddleOCR] 识别区域: {rect}')
            
            # 初始化OCR
            if not hasattr(self, '_paddle_ocr_engine'):
                log_info('[PaddleOCR] 初始化引擎 (检测到 Paddle 3.x/PIR 环境)...')
                # 针对 Paddle 3.x/Paddlex 3.x 的环境强制隔离
                os.environ["FLAGS_enable_pir_api"] = "0"       # 核心：关闭导致崩溃的 PIR 引擎
                os.environ["FLAGS_use_mkldnn"] = "0"
                os.environ["FLAGS_onednn_enabled"] = "0"
                os.environ["PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK"] = "True"
                
                # 按优先级尝试参数配置 (PaddleX 环境可能不支持某些旧参数，所以用 try_configs)
                configs = [
                    {"use_angle_cls": False, "lang": 'ch', "use_gpu": False, "show_log": False, "enable_mkldnn": False},
                    {"use_angle_cls": False, "lang": 'ch', "use_gpu": False},
                    {"lang": 'ch', "use_gpu": False},
                    {} # 最简初始化
                ]
                
                last_err = ""
                success = False
                for cfg in configs:
                    try:
                        self._paddle_ocr_engine = PaddleOCR(**cfg)
                        log_info(f'[PaddleOCR] 初始化成功 (参数: {list(cfg.keys())})')
                        success = True
                        break
                    except Exception as ex:
                        last_err = str(ex)
                        continue
                
                if not success:
                    log_error(f'[PaddleOCR] 所有尝试均失败. 最后报错: {last_err}')
                    # 尝试备选：直接通过 PaddleX 动态导入（如果环境允许）
                    try:
                        log_info('[PaddleOCR] 尝试通过 PaddleX 备选初始化...')
                        from paddlex import create_pipeline
                        self._paddle_ocr_engine = create_pipeline("OCR")
                        log_info('[PaddleOCR] 成功加载 PaddleX OCR 流水线')
                        success = True
                    except:
                        pass
                
                if not success:
                    return None
            
            # 裁剪图像
            arr = self._crop_region_to_numpy(rect)
            if arr is None:
                return None
            
            # 图像预处理 (Paddle专用版：保持灰度清晰度)
            arr = self.preprocess_image(arr, engine_type="paddleocr")
            
            # PaddleOCR识别 (针对 PaddleX/PaddleOCR 多版本兼容的健壮调用)
            recognized_text = ""
            try:
                # 1. 尝试不带参数的纯净调用 (兼容性最强)
                res = self._paddle_ocr_engine.ocr(arr)
                
                # 2. 如果返回为空，尝试带 cls 调用
                if res is None:
                    res = self._paddle_ocr_engine.ocr(arr, cls=True)
                
                # --- 开始解析结果 ---
                if res:
                    # 情况 A: 标准 PaddleOCR 格式 [ [[rect, (text, conf)], ...] ]
                    if isinstance(res, list) and len(res) > 0:
                        # 检查是否是 [[(text, conf)], ...] 这种 det=False 后的格式
                        if isinstance(res[0], list) and len(res[0]) > 0:
                            # 可能是标准格式
                            if isinstance(res[0][0], list) and len(res[0][0]) > 1:
                                texts = [line[1][0] for line in res[0] if line and len(line) > 1]
                                recognized_text = "".join(texts)
                            else:
                                # 可能是 det=False 格式
                                texts = [item[0] for item in res[0] if isinstance(item, (list, tuple))]
                                recognized_text = "".join(texts)
                        elif isinstance(res[0], tuple):
                            # 可能是 [('text', conf), ...]
                            recognized_text = "".join([item[0] for item in res if isinstance(item, tuple)])
                    
                    # 情况 B: PaddleX 结果对象 (PredictResult)
                    elif hasattr(res, "to_dict") or "PredictResult" in str(type(res)):
                        # 尝试从属性中提取文本
                        if hasattr(res, "ocr_text"):
                            recognized_text = "".join(res.ocr_text)
                        elif hasattr(res, "doc_res"): # 某些文档模式
                            recognized_text = "".join([x['text'] for x in res.doc_res])
            
            except Exception as e:
                # 如果 ocr() 彻底报错，尝试直接调用 predict (PaddleX 原生方式)
                try:
                    if hasattr(self._paddle_ocr_engine, 'predict'):
                        res = self._paddle_ocr_engine.predict(arr)
                        # 解析 PaddleX 典型输出
                        if isinstance(res, list) and len(res) > 0:
                            item = res[0]
                            if hasattr(item, 'doc_res'):
                                recognized_text = "".join([x.get('text', '') for x in item.doc_res])
                            elif hasattr(item, 'ocr_res'):
                                recognized_text = "".join([x.get('text', '') for x in item.ocr_res])
                except:
                    log_warn(f'[PaddleOCR] ⚠️ 多重兼容调用均失败: {e}')

                    return None

            if recognized_text:
                recognized_text = self._post_process_text(recognized_text)
                log_info(f'[PaddleOCR] ✓ 识别成功: {recognized_text}')
                return recognized_text
            
            log_error('[PaddleOCR] ✗ 未识别到文字')
            return None
                
        except ImportError:
            log_error('[PaddleOCR] ✗ 未安装（pip install paddleocr）')
            return None
        except Exception as e:
            log_error(f'[PaddleOCR] ✗ 识别失败: {e}')
            return None
    
    def _try_tesseract(self, rect):
        """尝试使用Tesseract识别"""
        try:
            import pytesseract
            from PIL import Image as PILImage
            import numpy as np
            
            log_info(f'[Tesseract] 识别区域: {rect}')
            
            # 配置Tesseract路径
            if self.tesseract_path and os.path.exists(self.tesseract_path):
                pytesseract.pytesseract.tesseract_cmd = self.tesseract_path
                log_info(f'[Tesseract] 使用配置路径: {self.tesseract_path}')
            
            # 裁剪图像
            arr = self._crop_region_to_numpy(rect)
            if arr is None:
                return None
            
            # 图像预处理 (Tesseract 强烈建议二值化)
            arr = self.preprocess_image(arr, engine_type="tesseract")
            
            # 转换为PIL Image
            pil_img = PILImage.fromarray(arr)
            
            # Tesseract识别
            text = pytesseract.image_to_string(pil_img, lang='chi_sim+eng')
            
            if text:
                processed_text = self._post_process_text(text)
                log_info(f'[Tesseract] ✓ 识别结果: {processed_text} (原始: {text.strip()})')
                return processed_text
            else:
                log_error('[Tesseract] ✗ 未识别到文字')
                return None
                
        except ImportError:
            log_error('[Tesseract] ✗ 未安装pytesseract（pip install pytesseract）')
            return None
        except Exception as e:
            log_error(f'[Tesseract] ✗ 识别失败: {e}')
            return None
    
    def _post_process_text(self, text):
        """后处理识别结果，精简及清洗噪音"""
        if not text:
            return ""
        
        # 1. 基础清理
        text = text.strip()
        
        # 2. 解决中文中间夹杂空格的问题 (针对 Tesseract)
        import re
        chinese_pattern = re.compile(r'[\u4e00-\u9fa5]')
        if chinese_pattern.search(text):
            # 如果包含中文，去掉明显的噪音符号
            text = text.replace("---", "").replace("- -", "")
            # 去掉文字间的空格，例如 "敬 据 同 步" -> "敬据同步"
            # 但保留中英文之间的空格（可选，这里为了准确率直接去全空格）
            text = "".join(text.split())
            
        # 3. 清理首尾非字符噪音 (针对 .-_ 等)
        text = re.sub(r'^[^a-zA-Z0-9\u4e00-\u9fa5]+', '', text)
        text = re.sub(r'[^a-zA-Z0-9\u4e00-\u9fa5]+$', '', text)
        
        # 4. 常见的 OCR 误识别字符及形近字纠正 (针对架构图常见词汇)
        replacements = {
            '敬据': '数据', '掌同': '同步', '惧改': '修改', '伯息': '信息',
            '痘信': '短信', '痛信': '短信', '算遮': '意愿', '葛廉': '意愿', 
            '詹遣': '意愿', '卵遣': '意愿', '探收': '接收', '算按': '模块', 
            '|': '', '１': '1', '２': '2', '３': '3', '４': '4', '５': '5',
            '６': '6', '７': '7', '８': '8', '９': '9', '０': '0',
        }
        for old, new in replacements.items():
            text = text.replace(old, new)
        
        return text.strip()
    
    def _crop_region_to_numpy(self, rect):
        """裁剪区域并转换为numpy数组 (添加适当外扩以提高OCR识别率)"""
        try:
            import numpy as np
            
            # 1. 区域微调：向外扩展 10px (PaddleOCR 需要更多留白)
            padding = 10
            x = max(0, rect.x() - padding)
            y = max(0, rect.y() - padding)
            w = min(self.original_pixmap.width() - x, rect.width() + padding * 2)
            h = min(self.original_pixmap.height() - y, rect.height() + padding * 2)
            
            padded_rect = QRect(x, y, w, h)
            
            # 2. 裁剪图像
            image = self.original_pixmap.toImage()
            cropped = image.copy(padded_rect)
            width, height = cropped.width(), cropped.height()
            
            if width < 5 or height < 5:
                return None
            
            # 3. 转换为RGB格式
            cropped = cropped.convertToFormat(QImage.Format_RGB888)
            bytes_per_line = cropped.bytesPerLine()
            ptr = cropped.bits()
            
            # 转换为numpy数组
            if hasattr(ptr, 'tobytes'):
                img_data = ptr.tobytes()
            else:
                img_data = ptr.asstring(bytes_per_line * height)
            
            arr = np.frombuffer(img_data, dtype=np.uint8)
            if bytes_per_line == width * 3:
                arr = arr.reshape((height, width, 3))
            else:
                arr = arr.reshape((height, bytes_per_line))
                arr = arr[:, :width*3].reshape((height, width, 3))
            
            return arr
            
        except Exception as e:
            log_error(f'[OCR] 图像裁剪转换失败: {e}')
            return None
    
    def paintEvent(self, event):
        """核心绘制逻辑 - 使用paintEvent提高性能和稳定性"""
        if not self.original_pixmap:
            super().paintEvent(event)
            return

        painter = QPainter(self)
        painter.setRenderHint(QPainter.SmoothPixmapTransform)

        # 1. 计算缩放后的图像大小
        scaled_width = int(self.original_pixmap.width() * self.scale_factor)
        scaled_height = int(self.original_pixmap.height() * self.scale_factor)
        
        # 2. 计算绘制位置 (居中 + 偏移)
        draw_x = (self.width() - scaled_width) // 2 + self.offset.x()
        draw_y = (self.height() - scaled_height) // 2 + self.offset.y()

        # 3. 绘制图像
        scaled_pixmap = self.original_pixmap.scaled(
            scaled_width, scaled_height,
            Qt.KeepAspectRatio,
            Qt.SmoothTransformation
        )
        painter.drawPixmap(draw_x, draw_y, scaled_pixmap)

        # 4. 绘制已有的标注
        for rect, num, name, color in self.annotations:
            scaled_rect = QRect(
                int(rect.x() * self.scale_factor) + draw_x,
                int(rect.y() * self.scale_factor) + draw_y,
                int(rect.width() * self.scale_factor),
                int(rect.height() * self.scale_factor)
            )
            
            # 画框
            pen = QPen(QColor(color), max(2, int(3 * self.scale_factor)))
            painter.setPen(pen)
            painter.drawRect(scaled_rect)
            
            # 标签
            label_text = f"{num} {name}"
            font_size = max(8, int(10 * self.scale_factor))
            painter.setFont(QFont("Microsoft YaHei", font_size, QFont.Bold))
            text_rect = painter.fontMetrics().boundingRect(label_text)
            
            bg_rect = QRect(
                scaled_rect.left(), 
                scaled_rect.top() - text_rect.height() - 4,
                text_rect.width() + 8,
                text_rect.height() + 4
            )
            painter.fillRect(bg_rect, QColor(color))
            painter.setPen(QColor("white"))
            painter.drawText(bg_rect, Qt.AlignCenter, label_text)

        # 5. 绘制当前正在框选的区域
        if self.is_drawing and self.current_rect:
            sel_rect = QRect(
                int(self.current_rect.x() * self.scale_factor) + draw_x,
                int(self.current_rect.y() * self.scale_factor) + draw_y,
                int(self.current_rect.width() * self.scale_factor),
                int(self.current_rect.height() * self.scale_factor)
            )
            pen = QPen(QColor("#10b981"), 2, Qt.DashLine)
            painter.setPen(pen)
            painter.drawRect(sel_rect)
            painter.fillRect(sel_rect, QColor(16, 185, 129, 30))

        painter.end()

    def update_display(self):
        """更新显示 - 现在只需触发更新重绘"""
        self.update()
    
    def clear_annotations(self):
        """清除所有标注"""
        self.annotations.clear()
        self.update_display()
    
    def remove_last_annotation(self):
        """移除最后一个标注"""
        if self.annotations:
            self.annotations.pop()
            self.update_display()
            return True
        return False
    
    def get_annotations(self):
        """获取所有标注 [(编号, 名称), ...]"""
        return [(num, name) for _, num, name, _ in self.annotations]


class ModuleInputDialog(QDialog):
    """模块信息输入对话框"""
    
    def __init__(self, parent=None, ocr_text=""):
        super().__init__(parent)
        self.setWindowTitle("输入模块信息")
        self.setModal(True)
        self.setFixedWidth(450)
        
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        
        # 说明
        tip = QLabel("💡 请输入标注模块的信息：")
        layout.addWidget(tip)
        
        # OCR识别提示（如果有）
        if ocr_text:
            ocr_tip = QLabel(f"🔍 OCR识别: {ocr_text[:80]}..." if len(ocr_text) > 80 else f"🔍 OCR识别: {ocr_text}")
            ocr_tip.setStyleSheet("color: #10b981; font-size: 11px; padding: 6px; background-color: rgba(16, 185, 129, 0.1); border-radius: 4px;")
            ocr_tip.setWordWrap(True)
            layout.addWidget(ocr_tip)
        else:
            # 没有识别到文字时的提示
            no_ocr_tip = QLabel("ℹ️ 未识别到文字，请手动输入（或安装PaddleOCR提升识别率）")
            no_ocr_tip.setStyleSheet("color: #9ca3af; font-size: 10px; padding: 4px;")
            no_ocr_tip.setWordWrap(True)
            layout.addWidget(no_ocr_tip)
        
        # 模块名称
        name_layout = QHBoxLayout()
        name_label = QLabel("模块名称:")
        name_label.setFixedWidth(80)
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("例如: 用户管理优化")
        # 如果有OCR结果，自动填充
        if ocr_text:
            self.name_input.setText(ocr_text.strip())
        name_layout.addWidget(name_label)
        name_layout.addWidget(self.name_input)
        layout.addLayout(name_layout)
        
        # 模块层级（替代模块编号）
        level_layout = QHBoxLayout()
        level_label = QLabel("模块层级:")
        level_label.setFixedWidth(80)
        self.level_combo = QComboBox()
        self.level_combo.addItem("一级模块", 1)
        self.level_combo.addItem("二级模块", 2)
        self.level_combo.addItem("三级模块", 3)
        self.level_combo.setCurrentIndex(1)  # 默认二级
        level_layout.addWidget(level_label)
        level_layout.addWidget(self.level_combo)
        layout.addLayout(level_layout)
        
        # 标注颜色（简化）
        color_layout = QHBoxLayout()
        color_label = QLabel("标注颜色:")
        color_label.setFixedWidth(80)
        self.color_combo = QComboBox()
        self.color_combo.addItem("� 橙色 (优化)", "#f59e0b")
        self.color_combo.addItem("🔴 红色 (新增)", "#ef4444")
        self.color_combo.addItem("🟢 绿色 (变更)", "#10b981")
        self.color_combo.addItem("🔵 蓝色 (已建)", "#3b82f6")
        color_layout.addWidget(color_label)
        color_layout.addWidget(self.color_combo)
        layout.addLayout(color_layout)
        
        # 按钮
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        
        cancel_btn = QPushButton("取消")
        cancel_btn.setFixedSize(80, 32)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)
        
        ok_btn = QPushButton("确定")
        ok_btn.setFixedSize(80, 32)
        ok_btn.setDefault(True)
        ok_btn.clicked.connect(self.accept)
        btn_layout.addWidget(ok_btn)
        
        layout.addLayout(btn_layout)
        
        # 焦点在模块名称输入框
        self.name_input.setFocus()
    
    def get_values(self):
        """获取输入值 (层级, 名称, 颜色)"""
        level = self.level_combo.currentData()
        name = self.name_input.text().strip()
        color = self.color_combo.currentData()
        
        # 根据层级生成编号占位符 (使用数字以便后端解析)
        level_prefix = {
            1: "1",      # 一级模块
            2: "1.1",    # 二级模块
            3: "1.1.1"   # 三级模块
        }
        num = level_prefix.get(level, "1.1")
        
        return (num, name, color)


class ImageAnnotatorDialog(QDialog):
    """图像标注主对话框"""
    
    def __init__(self, image_path, parent=None, initial_annotations=None):
        super().__init__(parent)
        self.image_path = image_path
        self.setWindowTitle("功能架构图标注 - 框选优化模块")
        self.resize(1400, 900)
        self.setModal(True)
        
        self.modules = []  # [(编号, 名称), ...]
        self._initial_annotations = initial_annotations or []
        
        self.init_ui()
        
        # 加载初始标注
        if self._initial_annotations:
            # self._initial_annotations 格式是 [(rect, num, name, color), ...]
            self.image_label.annotations = list(self._initial_annotations)
            for rect, num, name, color in self.image_label.annotations:
                # 更新列表
                item_text = f"{num}  {name}"
                item = QListWidgetItem(item_text)
                item.setToolTip(f"编号: {num}\n名称: {name}")
                item.setData(Qt.UserRole, rect)
                self.module_list.addItem(item)
            
            # 更新统计
            self.count_label.setText(f"共 {self.module_list.count()} 个模块")
            self.finish_btn.setEnabled(True)
            self.image_label.update_display()

    def init_ui(self):
        """初始化UI"""
        # 获取主题配置
        from extend.matcher_config import MatcherConfig
        config = MatcherConfig.load()
        is_dark = config.get("theme", {}).get("is_dark", False)
        
        # 配色方案
        bg_color = "#1f2937" if is_dark else "#f8fafc"
        card_bg = "#1f2937" if is_dark else "#ffffff"
        border_color = "#374151" if is_dark else "#e2e8f0"
        text_color = "#f3f4f6" if is_dark else "#1e293b"
        text_sec = "#9ca3af" if is_dark else "#64748b"
        input_bg = "#374151" if is_dark else "#ffffff"
        
        self.setStyleSheet(f"""
            QDialog {{ background-color: {bg_color}; color: {text_color}; }}
            QLabel {{ color: {text_color}; }}
            QPushButton {{ 
                background-color: {"#374151" if is_dark else "#e2e8f0"}; 
                color: {text_color};
                border: 1px solid {border_color};
                border-radius: 4px;
            }}
            QPushButton:hover {{ background-color: {"#4b5563" if is_dark else "#cbd5e1"}; }}
        """)

        main_layout = QHBoxLayout(self)
        main_layout.setSpacing(12)
        
        # 左侧 - 图像标注区域
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        
        # 工具栏
        toolbar = QFrame()
        toolbar.setStyleSheet(f"background-color: {card_bg}; border: 1px solid {border_color}; border-radius: 6px; padding: 8px;")
        toolbar_layout = QHBoxLayout(toolbar)
        
        tool_label = QLabel("🖱️ 左键框选 | 右键拖动 | 滚轮缩放")
        tool_label.setStyleSheet(f"color: {text_sec}; font-size: 12px;")
        toolbar_layout.addWidget(tool_label)
        
        # 缩放比例提示
        self.scale_label = QLabel("100%")
        self.scale_label.setStyleSheet("color: #10b981; font-size: 12px; padding: 4px 8px; background-color: rgba(16, 185, 129, 0.1); border-radius: 4px;")
        toolbar_layout.addWidget(self.scale_label)
        
        toolbar_layout.addStretch()
        
        # 缩放按钮
        zoom_in_btn = QPushButton("🔍+")
        zoom_in_btn.setFixedSize(50, 32)
        zoom_in_btn.setToolTip("放大（或使用鼠标滚轮）")
        zoom_in_btn.clicked.connect(self.zoom_in)
        toolbar_layout.addWidget(zoom_in_btn)
        
        zoom_out_btn = QPushButton("🔍-")
        zoom_out_btn.setFixedSize(50, 32)
        zoom_out_btn.setToolTip("缩小（或使用鼠标滚轮）")
        zoom_out_btn.clicked.connect(self.zoom_out)
        toolbar_layout.addWidget(zoom_out_btn)
        
        zoom_reset_btn = QPushButton("1:1")
        zoom_reset_btn.setFixedSize(50, 32)
        zoom_reset_btn.setToolTip("重置缩放")
        zoom_reset_btn.clicked.connect(self.zoom_reset)
        toolbar_layout.addWidget(zoom_reset_btn)
        
        undo_btn = QPushButton("↶ 撤销")
        undo_btn.setFixedSize(80, 32)
        undo_btn.clicked.connect(self.undo_last)
        toolbar_layout.addWidget(undo_btn)
        
        clear_btn = QPushButton("🗑️ 清空")
        clear_btn.setFixedSize(80, 32)
        clear_btn.clicked.connect(self.clear_all)
        toolbar_layout.addWidget(clear_btn)
        
        left_layout.addWidget(toolbar)
        
        # OCR配置区域
        ocr_config_box = QGroupBox("🔧 OCR配置")
        ocr_config_box.setStyleSheet(f"""
            QGroupBox {{
                background-color: {card_bg};
                border: 1px solid {border_color};
                border-radius: 6px;
                margin-top: 10px;
                padding-top: 15px;
                color: #10b981;
                font-weight: bold;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }}
        """)
        ocr_config_layout = QVBoxLayout(ocr_config_box)
        ocr_config_layout.setSpacing(8)
        
        # 第一行：OCR引擎选择
        engine_label = QLabel("启用的OCR引擎（按顺序尝试）:")
        engine_label.setStyleSheet(f"color: {text_sec}; font-size: 11px;")
        ocr_config_layout.addWidget(engine_label)
        
        engines_layout = QHBoxLayout()
        engines_layout.setSpacing(15)
        
        # 加载配置
        self.ocr_config = OCRConfig.load()
        # 确保初始值包含所有三个引擎，如果没有显式指定
        enabled_engines = self.ocr_config.get("enabled_engines")
        if not enabled_engines:
            enabled_engines = ["paddleocr", "easyocr", "tesseract"]
        
        self.easyocr_checkbox = QCheckBox("EasyOCR（推荐）")
        self.easyocr_checkbox.setChecked("easyocr" in enabled_engines)
        self.easyocr_checkbox.setStyleSheet(f"color: {text_color}; font-size: 11px;")
        engines_layout.addWidget(self.easyocr_checkbox)
        
        self.paddleocr_checkbox = QCheckBox("PaddleOCR")
        self.paddleocr_checkbox.setChecked("paddleocr" in enabled_engines)
        self.paddleocr_checkbox.setStyleSheet(f"color: {text_color}; font-size: 11px;")
        engines_layout.addWidget(self.paddleocr_checkbox)
        
        self.tesseract_checkbox = QCheckBox("Tesseract")
        self.tesseract_checkbox.setChecked("tesseract" in enabled_engines)
        self.tesseract_checkbox.setStyleSheet(f"color: {text_color}; font-size: 11px;")
        engines_layout.addWidget(self.tesseract_checkbox)
        
        engines_layout.addStretch()
        ocr_config_layout.addLayout(engines_layout)
        
        # 第二行：Tesseract路径
        tesseract_path_layout = QHBoxLayout()
        tesseract_label = QLabel("Tesseract路径:")
        tesseract_label.setStyleSheet(f"color: {text_sec}; font-size: 11px;")
        tesseract_path_layout.addWidget(tesseract_label)
        
        self.tesseract_path_input = QLineEdit()
        self.tesseract_path_input.setText(self.ocr_config.get("tesseract_path", ""))
        self.tesseract_path_input.setPlaceholderText("C:\\Program Files\\Tesseract-OCR\\tesseract.exe")
        self.tesseract_path_input.setStyleSheet(f"""
            QLineEdit {{
                background-color: {input_bg};
                color: {text_color};
                border: 1px solid {border_color};
                border-radius: 4px;
                padding: 4px 8px;
                font-size: 11px;
            }}
        """)
        tesseract_path_layout.addWidget(self.tesseract_path_input, 1)
        
        browse_btn = QPushButton("浏览...")
        browse_btn.setFixedSize(60, 26)
        browse_btn.setStyleSheet("font-size: 11px;")
        browse_btn.clicked.connect(self.browse_tesseract_path)
        tesseract_path_layout.addWidget(browse_btn)
        
        ocr_config_layout.addLayout(tesseract_path_layout)
        
        # 第三行：图像增强+保存配置
        bottom_row = QHBoxLayout()
        self.enhance_checkbox = QCheckBox("启用图像增强（提高识别率）")
        self.enhance_checkbox.setChecked(self.ocr_config.get("enable_image_enhance", True))
        self.enhance_checkbox.setStyleSheet(f"color: {text_color}; font-size: 11px;")
        self.enhance_checkbox.setToolTip("对图像进行灰度化、增强对比度、降噪、二值化等处理")
        bottom_row.addWidget(self.enhance_checkbox)
        
        bottom_row.addStretch()
        
        save_config_btn = QPushButton("💾 保存配置")
        save_config_btn.setFixedSize(90, 26)
        save_config_btn.setStyleSheet("""
            QPushButton {
                background-color: #10b981;
                color: white;
                border-radius: 4px;
                font-size: 11px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #059669;
            }
        """)
        save_config_btn.clicked.connect(self.save_ocr_config)
        bottom_row.addWidget(save_config_btn)
        
        ocr_config_layout.addLayout(bottom_row)

        # 实时同步状态到 label
        self.easyocr_checkbox.stateChanged.connect(self.sync_engines_to_label)
        self.paddleocr_checkbox.stateChanged.connect(self.sync_engines_to_label)
        self.tesseract_checkbox.stateChanged.connect(self.sync_engines_to_label)
        self.enhance_checkbox.stateChanged.connect(self.sync_engines_to_label)
        
        left_layout.addWidget(ocr_config_box)
        
        # 图像显示区域（可滚动）
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("border: 1px solid #374151; border-radius: 6px;")
        
        self.image_label = AnnotationLabel()
        self.image_label.annotation_added.connect(self.on_annotation_added)
        
        # 将配置应用到 image_label
        self.image_label.enabled_engines = enabled_engines
        self.image_label.tesseract_path = self.ocr_config.get("tesseract_path", "")
        self.image_label.enable_image_enhance = self.ocr_config.get("enable_image_enhance", True)
        
        if not self.image_label.load_image(self.image_path):
            QMessageBox.warning(self, "错误", "无法加载图片")
            self.reject()
            return
        
        scroll.setWidget(self.image_label)
        left_layout.addWidget(scroll)
        
        # 初始化时同步一次配置
        self.sync_engines_to_label()
        
        main_layout.addWidget(left_widget, stretch=7)
        
        # 右侧 - 标注列表
        right_widget = QWidget()
        right_widget.setFixedWidth(350)
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        
        # 标题
        list_title = QLabel("📋 已标注模块列表")
        list_title.setStyleSheet("font-size: 14px; font-weight: bold; margin-bottom: 8px;")
        right_layout.addWidget(list_title)
        
        # 统计
        self.count_label = QLabel("共 0 个模块")
        self.count_label.setStyleSheet("color: #9ca3af; font-size: 12px; margin-bottom: 8px;")
        right_layout.addWidget(self.count_label)
        
        # 列表
        self.module_list = QListWidget()
        self.module_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.module_list.customContextMenuRequested.connect(self.show_list_context_menu)
        self.module_list.setStyleSheet(f"""
            QListWidget {{
                border: 1px solid {border_color};
                border-radius: 6px;
                background-color: {card_bg};
                color: {text_color};
                padding: 4px;
            }}
            QListWidget::item {{
                padding: 8px;
                border-radius: 4px;
                margin: 2px;
                color: {text_color};
            }}
            QListWidget::item:hover {{
                background-color: {input_bg};
            }}
            QListWidget::item:selected {{
                background-color: #3b82f6;
                color: white;
            }}
        """)
        right_layout.addWidget(self.module_list)
        
        # 帮助文本
        help_text = QLabel(
            "💡 使用说明:\n"
            "1. 鼠标滚轮缩放图片，便于查看细节\n"
            "2. 右键拖动平移图片（放大后使用）\n"
            "3. 左键拖动框选模块区域\n"
            "4. 系统自动OCR识别（可修改）\n"
            "5. 选择模块层级和颜色\n"
            "6. 标注完成后点击'完成标注'"
        )
        help_text.setStyleSheet(f"""
            color: {text_sec};
            font-size: 11px;
            background-color: {card_bg};
            border: 1px solid {border_color};
            border-radius: 6px;
            padding: 12px;
            margin-top: 12px;
        """)
        help_text.setWordWrap(True)
        right_layout.addWidget(help_text)
        
        # 底部按钮
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(8)
        
        cancel_btn = QPushButton("取消")
        cancel_btn.setFixedHeight(40)
        cancel_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: transparent;
                color: {text_color};
                border: 1px solid {border_color};
            }}
            QPushButton:hover {{
                background-color: {input_bg};
            }}
        """)
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)
        
        self.finish_btn = QPushButton("✓ 完成标注")
        self.finish_btn.setFixedHeight(40)
        self.finish_btn.setEnabled(False)
        self.finish_btn.clicked.connect(self.accept)
        self.finish_btn.setStyleSheet("""
            QPushButton {
                background-color: #10b981;
                color: white;
                border: none;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #059669;
            }
            QPushButton:disabled {
                background-color: #374151;
                color: #6b7280;
            }
        """)
        btn_layout.addWidget(self.finish_btn)
        
        right_layout.addLayout(btn_layout)
        
        main_layout.addWidget(right_widget, stretch=3)

    def browse_tesseract_path(self):
        """浏览选择Tesseract可执行文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "选择Tesseract可执行文件",
            "C:\\Program Files",
            "可执行文件 (*.exe);;所有文件 (*.*)"
        )
        
        if file_path:
            self.tesseract_path_input.setText(file_path)
    
    def save_ocr_config(self):
        """保存OCR配置"""
        # 收集启用的引擎
        enabled_engines = []
        if self.easyocr_checkbox.isChecked():
            enabled_engines.append("easyocr")
        if self.paddleocr_checkbox.isChecked():
            enabled_engines.append("paddleocr")
        if self.tesseract_checkbox.isChecked():
            enabled_engines.append("tesseract")
        
        # 构建配置
        config = {
            "enabled_engines": enabled_engines,
            "tesseract_path": self.tesseract_path_input.text().strip(),
            "enable_image_enhance": self.enhance_checkbox.isChecked()
        }
        
        # 保存到文件
        OCRConfig.save(config)
        
        # 更新image_label的配置
        self.image_label.enabled_engines = enabled_engines
        self.image_label.tesseract_path = config["tesseract_path"]
        self.image_label.enable_image_enhance = config["enable_image_enhance"]
        
        QMessageBox.information(self, "成功", "OCR配置已保存！\n下次打开时将自动应用此配置")

    def sync_engines_to_label(self):
        """实时同步勾选的引擎到标注组件 (无需点击保存即可生效)"""
        enabled_engines = []
        # 按照用户预想的优先级尝试
        if self.paddleocr_checkbox.isChecked():
            enabled_engines.append("paddleocr")
        if self.easyocr_checkbox.isChecked():
            enabled_engines.append("easyocr")
        if self.tesseract_checkbox.isChecked():
            enabled_engines.append("tesseract")
        
        self.image_label.enabled_engines = enabled_engines
        self.image_label.enable_image_enhance = self.enhance_checkbox.isChecked()
        self.image_label.tesseract_path = self.tesseract_path_input.text().strip()
        log_info(f'[OCR] 实时同步引擎配置: {enabled_engines}')
    
    def show_list_context_menu(self, pos):
        """显示列表右键菜单"""
        item = self.module_list.itemAt(pos)
        if not item:
            return
            
        from PySide6.QtWidgets import QMenu
        from PySide6.QtGui import QAction
        
        menu = QMenu(self)
        edit_action = QAction("✏️ 修改标注", self)
        edit_action.triggered.connect(lambda: self.edit_module_from_list(item))
        menu.addAction(edit_action)
        
        delete_action = QAction("🗑️ 删除标注", self)
        delete_action.triggered.connect(lambda: self.delete_module_from_list(item))
        menu.addAction(delete_action)
        
        menu.exec(self.module_list.mapToGlobal(pos))

    def edit_module_from_list(self, item):
        """从列表项触发修改"""
        rect = item.data(Qt.UserRole)
        if rect:
            # 查找对应的标注数据
            for r, num, name, color in self.image_label.annotations:
                if r == rect:
                    self.image_label.show_module_input_dialog(r, (num, name, color))
                    break

    def delete_module_from_list(self, item):
        """从列表项触发删除"""
        rect = item.data(Qt.UserRole)
        if rect:
            # 从标注中移除
            self.image_label.annotations = [ann for ann in self.image_label.annotations if ann[0] != rect]
            self.image_label.update_display()
            # 从列表中移除
            self.module_list.takeItem(self.module_list.row(item))
            self.count_label.setText(f"共 {self.module_list.count()} 个模块")
            if self.module_list.count() == 0:
                self.finish_btn.setEnabled(False)

    def on_annotation_added(self, rect, num, name):
        """标注添加回调"""
        # 更新列表
        item_text = f"{num}  {name}"
        
        # 查找是否已经存在（通过 rect 判断是否为编辑）
        found = False
        for i in range(self.module_list.count()):
            it = self.module_list.item(i)
            if it.data(Qt.UserRole) == rect:
                it.setText(item_text)
                found = True
                break
        
        if not found:
            item = QListWidgetItem(item_text)
            item.setToolTip(f"编号: {num}\n名称: {name}")
            item.setData(Qt.UserRole, rect)
            self.module_list.addItem(item)
        
        # 更新统计
        self.count_label.setText(f"共 {self.module_list.count()} 个模块")
        
        # 启用完成按钮
        self.finish_btn.setEnabled(True)
    
    def undo_last(self):
        """撤销最后一个标注"""
        if self.image_label.remove_last_annotation():
            if self.module_list.count() > 0:
                self.module_list.takeItem(self.module_list.count() - 1)
                self.count_label.setText(f"共 {self.module_list.count()} 个模块")
                
                if self.module_list.count() == 0:
                    self.finish_btn.setEnabled(False)
    
    def zoom_in(self):
        """放大图片"""
        if hasattr(self.image_label, 'scale_factor'):
            new_scale = min(self.image_label.scale_factor * 1.2, self.image_label.max_scale)
            self.image_label.scale_factor = new_scale
            self.image_label.update_display()
            self.show_scale_tip(int(new_scale * 100))
    
    def zoom_out(self):
        """缩小图片"""
        if hasattr(self.image_label, 'scale_factor'):
            new_scale = max(self.image_label.scale_factor * 0.8, self.image_label.min_scale)
            self.image_label.scale_factor = new_scale
            self.image_label.update_display()
            self.show_scale_tip(int(new_scale * 100))
    
    def zoom_reset(self):
        """重置缩放"""
        if hasattr(self.image_label, 'scale_factor'):
            self.image_label.scale_factor = 1.0
            self.image_label.offset = QPoint(0, 0)  # 重置偏移
            self.image_label.update_display()
            self.show_scale_tip(100)
    
    def show_scale_tip(self, percentage):
        """显示缩放比例提示"""
        self.scale_label.setText(f"{percentage}%")
    
    def clear_all(self):
        """清空所有标注"""
        if not self.modules:
            return
        
        reply = QMessageBox.question(
            self, 
            "确认清空", 
            "确定要清空所有标注吗？",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if reply == QMessageBox.Yes:
            self.image_label.clear_annotations()
            self.modules.clear()
            self.module_list.clear()
            self.count_label.setText("共 0 个模块")
            self.finish_btn.setEnabled(False)
    
    def get_modules(self):
        """获取标注的完整数据 [(rect, num, name, color), ...]"""
        return self.image_label.annotations
