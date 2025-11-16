#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
论文格式自动转换工具 - Flask Web应用
"""
from flask import Flask, request, render_template, send_file, jsonify
from werkzeug.utils import secure_filename
import os
from datetime import datetime
from format_converter import PaperFormatConverter

app = Flask(__name__)

# 配置
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
ALLOWED_EXTENSIONS = {'docx'}
MAX_FILE_SIZE = 10 * 1024 * 1024  # 10MB

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_FILE_SIZE

# 确保目录存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def allowed_file(filename):
    """检查文件扩展名是否允许"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    """主页"""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """处理文件上传和格式转换"""
    try:
        # 检查是否有文件
        if 'file' not in request.files:
            return jsonify({'error': '没有上传文件'}), 400

        file = request.files['file']

        # 检查文件名
        if file.filename == '':
            return jsonify({'error': '文件名为空'}), 400

        # 检查文件类型
        if not allowed_file(file.filename):
            return jsonify({'error': '只支持 .docx 格式的文件'}), 400

        # 保存上传的文件
        filename = secure_filename(file.filename)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        original_path = os.path.join(app.config['UPLOAD_FOLDER'], f"{timestamp}_{filename}")
        file.save(original_path)

        # 生成输出文件名
        output_filename = f"{timestamp}_formatted_{filename}"
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

        # 转换格式
        converter = PaperFormatConverter()
        success, message = converter.convert(original_path, output_path)

        if success:
            return jsonify({
                'success': True,
                'message': message,
                'download_url': f'/download/{output_filename}'
            })
        else:
            return jsonify({'error': message}), 500

    except Exception as e:
        return jsonify({'error': f'处理失败: {str(e)}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    """下载转换后的文件"""
    try:
        file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
        if os.path.exists(file_path):
            return send_file(
                file_path,
                as_attachment=True,
                download_name=filename
            )
        else:
            return jsonify({'error': '文件不存在'}), 404
    except Exception as e:
        return jsonify({'error': f'下载失败: {str(e)}'}), 500

@app.route('/format-info')
def format_info():
    """返回格式要求信息"""
    info = {
        'page_setup': {
            '纸张大小': 'A4 (21.00 x 29.70 cm)',
            '上边距': '2.5 cm',
            '下边距': '2.5 cm',
            '左边距': '3.0 cm',
            '右边距': '3.0 cm'
        },
        'styles': {
            '一级标题': {
                '字号': '16 pt',
                '对齐': '居中',
                '段前': '6 pt',
                '段后': '6 pt'
            },
            '二级标题': {
                '字体': '宋体',
                '首行缩进': '0.99 cm'
            },
            '三级标题': {
                '字号': '12 pt'
            },
            '正文': {
                '字体': 'Times New Roman / 宋体',
                '字号': '16 pt',
                '首行缩进': '0.85 cm',
                '行距': '固定值'
            }
        }
    }
    return jsonify(info)

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
