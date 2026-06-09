# -*- coding: utf-8 -*-
"""
教师作业批改助手 - Flask 后端
Teacher Homework Grading Assistant - Flask Backend
"""

import os
import json
import time
import requests
from flask import Flask, render_template, request, jsonify, Response, stream_with_context, make_response
from flask_cors import CORS
from docx import Document
import PyPDF2
import io

app = Flask(__name__)
app.json.ensure_ascii = True  # 强制 JSON 转义非 ASCII 字符，避免 latin-1 编码问题
CORS(app)


def safe_jsonify(data, status=200):
    """安全的 JSON 响应 —— 手动序列化，确保 UTF-8，杜绝 latin-1"""
    body = json.dumps(data, ensure_ascii=True, sort_keys=False)
    resp = make_response(body, status)
    resp.headers["Content-Type"] = "application/json; charset=utf-8"
    return resp


# ============================================================
# 默认配置 - 用户可在前端修改
# ============================================================

DEFAULT_CONFIG = {
    "api_providers": [
        {
            "id": "siliconflow",
            "name": "SiliconFlow (硅基流动)",
            "base_url": "https://api.siliconflow.cn/v1",
            "models": [
                {"id": "Qwen/Qwen2.5-7B-Instruct", "name": "Qwen2.5-7B (免费)", "free": True},
                {"id": "Qwen/Qwen2.5-14B-Instruct", "name": "Qwen2.5-14B", "free": False},
                {"id": "deepseek-ai/DeepSeek-V3", "name": "DeepSeek-V3", "free": False},
                {"id": "Qwen/Qwen3-8B", "name": "Qwen3-8B (免费)", "free": True},
            ]
        },
        {
            "id": "deepseek",
            "name": "DeepSeek 官方",
            "base_url": "https://api.deepseek.com",
            "models": [
                {"id": "deepseek-chat", "name": "DeepSeek-V3", "free": False},
                {"id": "deepseek-reasoner", "name": "DeepSeek-R1", "free": False},
            ]
        },
        {
            "id": "moonshot",
            "name": "Moonshot (Kimi 月之暗面)",
            "base_url": "https://api.moonshot.cn/v1",
            "default_api_key": "sk-ROSBzcRzxXZRW17vIRLmXsVNzEV9Yq0H74YEg8lsTTR87L2g",
            "models": [
                {"id": "moonshot-v1-8k", "name": "Moonshot V1-8K (便宜)", "free": False},
                {"id": "moonshot-v1-32k", "name": "Moonshot V1-32K", "free": False},
                {"id": "moonshot-v1-128k", "name": "Moonshot V1-128K (长文本)", "free": False},
                {"id": "kimi-k2", "name": "Kimi K2", "free": False},
                {"id": "kimi-k2.5", "name": "Kimi K2.5 (最新旗舰)", "free": False},
            ]
        },
        {
            "id": "openai_compatible",
            "name": "自定义 OpenAI 兼容 API",
            "base_url": "",
            "models": []
        }
    ],
    "courses": [
        {"id": "chinese", "name": "语文", "icon": "📝", "subjects": ["阅读理解", "作文", "古诗词鉴赏", "文言文翻译", "现代文阅读"]},
        {"id": "math", "name": "数学", "icon": "🔢", "subjects": ["代数", "几何", "概率统计", "函数", "方程与不等式"]},
        {"id": "english", "name": "英语", "icon": "🔤", "subjects": ["阅读理解", "写作", "完形填空", "语法填空", "翻译"]},
        {"id": "physics", "name": "物理", "icon": "⚡", "subjects": ["力学", "电磁学", "热学", "光学", "原子物理"]},
        {"id": "chemistry", "name": "化学", "icon": "🧪", "subjects": ["有机化学", "无机化学", "化学反应原理", "实验题", "化学方程式"]},
        {"id": "biology", "name": "生物", "icon": "🧬", "subjects": ["细胞生物学", "遗传学", "生态学", "人体生理", "进化论"]},
        {"id": "history", "name": "历史", "icon": "📜", "subjects": ["中国古代史", "中国近代史", "世界史", "材料分析", "论述题"]},
        {"id": "geography", "name": "地理", "icon": "🌍", "subjects": ["自然地理", "人文地理", "区域地理", "地图分析", "综合题"]},
        {"id": "politics", "name": "政治", "icon": "⚖️", "subjects": ["经济生活", "政治生活", "文化生活", "哲学", "时政分析"]},
        {"id": "programming", "name": "编程/信息", "icon": "💻", "subjects": ["Python", "C/C++", "算法", "数据结构", "程序填空"]},
        {"id": "custom", "name": "自定义课程", "icon": "📚", "subjects": []},
    ]
}

# ============================================================
# 路由
# ============================================================

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/config", methods=["GET"])
def get_config():
    """获取前端配置"""
    return jsonify(DEFAULT_CONFIG)

@app.route("/manual")
def product_manual():
    """产品说明书"""
    return render_template("manual.html")


@app.route("/api/grade", methods=["POST"])
def grade_homework():
    """批改作业核心接口"""
    try:
        data = request.get_json()

        # 提取参数
        homework_text = data.get("homework_text", "").strip()
        course = data.get("course", "")
        subject = data.get("subject", "")
        api_provider = data.get("api_provider", "siliconflow")
        api_key = data.get("api_key", "").strip()
        model = data.get("model", "")
        custom_url = data.get("custom_url", "").strip()
        extra_instructions = data.get("extra_instructions", "").strip()
        grading_criteria = data.get("grading_criteria", "standard")

        if not homework_text:
            return safe_jsonify({"error": "请输入作业内容"}, 400)
        if not api_key:
            return safe_jsonify({"error": "请输入 API Key"}, 400)
        if not course:
            return safe_jsonify({"error": "请选择课程"}, 400)

        # 构建批改提示词
        system_prompt = build_system_prompt(course, subject, grading_criteria, extra_instructions)
        user_prompt = build_user_prompt(homework_text, course, subject)

        # 调用 AI API
        result = call_ai_api(
            api_provider=api_provider,
            api_key=api_key,
            model=model,
            system_prompt=system_prompt,
            user_prompt=user_prompt,
            custom_url=custom_url
        )

        return safe_jsonify({
            "success": True,
            "result": result,
            "metadata": {
                "course": course,
                "subject": subject,
                "model": model,
                "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                "word_count": len(homework_text),
            }
        })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return safe_jsonify({"error": f"批改失败: {str(e)}"}, 500)


@app.route("/api/grade/stream", methods=["POST"])
def grade_homework_stream():
    """批改作业 - 流式输出（SSE）"""
    try:
        data = request.get_json()
        homework_text = data.get("homework_text", "").strip()
        course = data.get("course", "")
        subject = data.get("subject", "")
        api_provider = data.get("api_provider", "siliconflow")
        api_key = data.get("api_key", "").strip()
        model = data.get("model", "")
        custom_url = data.get("custom_url", "").strip()
        extra_instructions = data.get("extra_instructions", "").strip()
        grading_criteria = data.get("grading_criteria", "standard")

        if not homework_text:
            return safe_jsonify({"error": "请输入作业内容"}, 400)
        if not api_key:
            return safe_jsonify({"error": "请输入 API Key"}, 400)
        if not course:
            return safe_jsonify({"error": "请选择课程"}, 400)

        system_prompt = build_system_prompt(course, subject, grading_criteria, extra_instructions)
        user_prompt = build_user_prompt(homework_text, course, subject)

        def generate():
            try:
                for chunk in call_ai_api_stream(
                    api_provider=api_provider,
                    api_key=api_key,
                    model=model,
                    system_prompt=system_prompt,
                    user_prompt=user_prompt,
                    custom_url=custom_url
                ):
                    yield f"data: {json.dumps({'content': chunk}, ensure_ascii=True)}\n\n"
                # 发送元数据
                meta = json.dumps({
                    "done": True,
                    "metadata": {
                        "course": course,
                        "subject": subject,
                        "model": model,
                        "timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                        "word_count": len(homework_text),
                    }
                }, ensure_ascii=True)
                yield f"data: {meta}\n\n"
            except Exception as e:
                yield f"data: {json.dumps({'error': str(e)}, ensure_ascii=True)}\n\n"

        return Response(
            stream_with_context(generate()),
            mimetype="text/event-stream; charset=utf-8",
            headers={
                "Cache-Control": "no-cache",
                "X-Accel-Buffering": "no",
            }
        )
    except Exception as e:
        return safe_jsonify({"error": f"批改失败: {str(e)}"}, 500)


@app.route("/api/upload", methods=["POST"])
def upload_file():
    """上传文件并提取文本"""
    try:
        if "file" not in request.files:
            return safe_jsonify({"error": "未选择文件"}, 400)

        file = request.files["file"]
        if file.filename == "":
            return safe_jsonify({"error": "未选择文件"}, 400)

        filename = file.filename.lower()
        text = ""

        if filename.endswith(".txt"):
            text = file.read().decode("utf-8", errors="ignore")
        elif filename.endswith(".docx"):
            doc = Document(io.BytesIO(file.read()))
            text = "\n".join([para.text for para in doc.paragraphs if para.text.strip()])
        elif filename.endswith(".pdf"):
            reader = PyPDF2.PdfReader(io.BytesIO(file.read()))
            pages = []
            for page in reader.pages:
                page_text = page.extract_text()
                if page_text:
                    pages.append(page_text)
            text = "\n".join(pages)
        elif filename.endswith(".md"):
            text = file.read().decode("utf-8", errors="ignore")
        else:
            return safe_jsonify({"error": f"不支持的文件格式: {filename}。支持 .txt, .docx, .pdf, .md"}, 400)

        if not text.strip():
            return safe_jsonify({"error": "文件内容为空，无法提取文字"}, 400)

        return safe_jsonify({
            "success": True,
            "text": text.strip(),
            "filename": file.filename,
            "char_count": len(text.strip())
        })

    except Exception as e:
        return safe_jsonify({"error": f"文件解析失败: {str(e)}"}, 500)


# ============================================================
# 提示词构建
# ============================================================

def build_system_prompt(course, subject, grading_criteria, extra_instructions):
    """构建系统提示词"""

    criteria_map = {
        "strict": "严格标准：按照考试评分标准严格批改，不放过任何错误，扣分明确。",
        "standard": "标准标准：按照常规教学要求批改，指出错误并给改进建议。",
        "lenient": "宽松标准：以鼓励为主，重点肯定做得好的地方，温和指出需要改进的地方。",
    }

    criteria_desc = criteria_map.get(grading_criteria, criteria_map["standard"])

    system = f"""你是一位经验丰富的{course}教师，擅长批改{subject}相关的作业。

## 批改标准
{criteria_desc}

## 你的任务
请认真阅读学生提交的作业，完成以下工作：

1. **评分**：根据作业质量给出百分制评分（0-100分）
2. **总体评价**：用2-3句话概括作业的整体表现
3. **详细批改**：
   - 逐题或逐段批改（如适用）
   - 标注正确的内容（用 ✅ 标记）
   - 标注错误的内容（用 ❌ 标记）
   - 标注需要改进的地方（用 ⚠️ 标记）
4. **优点**：列出作业中做得好的地方（至少1条）
5. **不足与建议**：列出需要改进的地方，并给出具体改进建议
6. **鼓励语**：给学生一句鼓励的话

## 输出格式
请使用以下 Markdown 格式输出，便于前端渲染。

**格式要求**：
- 数学表达式用纯文本书写，不要用 $...$ 或 $$...$$ 等 LaTeX 公式
- 分数写成 1/2，不要写 \frac{1}{2}
- 平方用 ²、立方用 ³、乘号用 ×、除号用 ÷、根号用 √
- 下标直接写 x1、x2，不要写 x_1、x_2
- 整体风格适合学生阅读，不要出现编程符号

### 📊 评分：XX / 100 分

### 📋 总体评价
（总体评价内容）

### 📝 详细批改
（逐题/逐段批改内容，正确用 ✅、错误用 ❌、需改进用 ⚠️ 标记）

### ✨ 优点
- 优点1
- 优点2

### ⚠️ 不足与建议
- 不足1 → 改进建议
- 不足2 → 改进建议

### 💬 教师寄语
（鼓励的话）
"""

    if extra_instructions:
        system += f"\n## 额外要求\n{extra_instructions}\n"

    return system


def build_user_prompt(homework_text, course, subject):
    """构建用户提示词"""
    return f"""以下是学生提交的{course}（{subject}）作业：

---
{homework_text}
---

请按照要求认真批改这份作业。"""


# ============================================================
# AI API 调用
# ============================================================

def call_ai_api(api_provider, api_key, model, system_prompt, user_prompt, custom_url=""):
    """调用 AI API"""

    # 确定 API URL
    url_map = {
        "siliconflow": "https://api.siliconflow.cn/v1/chat/completions",
        "deepseek": "https://api.deepseek.com/v1/chat/completions",
        "moonshot": "https://api.moonshot.cn/v1/chat/completions",
        "openai_compatible": f"{custom_url.rstrip('/')}/chat/completions" if custom_url else "",
    }

    url = url_map.get(api_provider, "")
    if not url:
        raise ValueError("无效的 API 提供商或缺少自定义 URL")

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json; charset=utf-8"
    }

    # 推理模型（如 DeepSeek-R1）只接受 temperature=1
    # 普通模型用 0.3 以获得稳定的批改结果
    reasoning_models = ["deepseek-reasoner", "deepseek-r1", "kimi-k2.5"]
    is_reasoning = any(m in model.lower() for m in reasoning_models)

    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": 1.0 if is_reasoning else 0.3,
        "max_tokens": 2048,
    }

    # 手动序列化 JSON，使用 ensure_ascii=True 确保所有非 ASCII 字符被转义
    # 避免 Windows 下 latin-1 编码问题
    body = json.dumps(payload, ensure_ascii=True).encode("utf-8")
    response = requests.post(url, headers=headers, data=body, timeout=120)

    if response.status_code != 200:
        error_msg = f"API 调用失败 (HTTP {response.status_code})"
        try:
            # 用 response.content 获取原始字节，手动按 UTF-8 解码
            raw = response.content
            err = json.loads(raw.decode("utf-8"))
            error_msg += f": {err.get('error', {}).get('message', raw.decode('utf-8', errors='replace')[:200])}"
        except:
            error_msg += f": {response.content.decode('utf-8', errors='replace')[:200]}"
        raise Exception(error_msg)

    # 手动解码响应，避免 requests.text 的编码猜测问题
    raw = response.content.decode("utf-8")
    result = json.loads(raw)
    content = result.get("choices", [{}])[0].get("message", {}).get("content", "")

    if not content:
        raise Exception("AI 返回内容为空，请检查模型和 API Key 是否正确")

    return content


def call_ai_api_stream(api_provider, api_key, model, system_prompt, user_prompt, custom_url=""):
    """调用 AI API - 流式输出"""

    url_map = {
        "siliconflow": "https://api.siliconflow.cn/v1/chat/completions",
        "deepseek": "https://api.deepseek.com/v1/chat/completions",
        "moonshot": "https://api.moonshot.cn/v1/chat/completions",
        "openai_compatible": f"{custom_url.rstrip('/')}/chat/completions" if custom_url else "",
    }

    url = url_map.get(api_provider, "")
    if not url:
        raise ValueError("无效的 API 提供商或缺少自定义 URL")

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json; charset=utf-8"
    }

    reasoning_models = ["deepseek-reasoner", "deepseek-r1", "kimi-k2.5"]
    is_reasoning = any(m in model.lower() for m in reasoning_models)

    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt}
        ],
        "temperature": 1.0 if is_reasoning else 0.3,
        "max_tokens": 2048,
        "stream": True,
    }

    body = json.dumps(payload, ensure_ascii=True).encode("utf-8")
    response = requests.post(url, headers=headers, data=body, timeout=180, stream=True)

    if response.status_code != 200:
        try:
            err = json.loads(response.content.decode("utf-8"))
            raise Exception(f"API 调用失败 (HTTP {response.status_code}): {err.get('error', {}).get('message', '')}")
        except json.JSONDecodeError:
            raise Exception(f"API 调用失败 (HTTP {response.status_code})")

    # 逐行解析 SSE 流
    for line in response.iter_lines(decode_unicode=True):
        if not line or not line.startswith("data: "):
            continue
        data_str = line[6:]  # 去掉 "data: " 前缀
        if data_str == "[DONE]":
            break
        try:
            chunk = json.loads(data_str)
            delta = chunk.get("choices", [{}])[0].get("delta", {})
            content = delta.get("content", "")
            if content:
                yield content
        except json.JSONDecodeError:
            continue


# ============================================================
# 启动
# ============================================================

if __name__ == "__main__":
    import sys
    import io
    # Fix Windows console encoding for emoji output
    if sys.platform == "win32":
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

    print("\n" + "=" * 60)
    print("  [GRAD] 教师作业批改助手 v1.1")
    print("  Teacher Homework Grading Assistant")
    print("=" * 60)
    print("  访问地址: http://localhost:5000")
    print("=" * 60 + "\n")

    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
