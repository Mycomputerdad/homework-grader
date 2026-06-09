# -*- coding: utf-8 -*-
"""生成桌面 Word 产品说明书"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
import datetime

doc = Document()

# 页面设置
for section in doc.sections:
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

# 默认样式
style = doc.styles['Normal']
style.font.name = '宋体'
style.font.size = Pt(10.5)
style.font.color.rgb = RGBColor(0, 0, 0)
style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

BLACK = RGBColor(0, 0, 0)


def set_font(run, size=Pt(10.5), bold=False, underline=False):
    run.font.name = '宋体'
    run.font.size = size
    run.font.bold = bold
    run.font.color.rgb = BLACK
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    if underline:
        run.underline = True


def add_title(text, level=0):
    p = doc.add_paragraph()
    if level == 0:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(text)
        set_font(run, Pt(22), bold=True)
    elif level == 1:
        p.paragraph_format.space_before = Pt(18)
        run = p.add_run(text)
        set_font(run, Pt(15), bold=True)
    elif level == 2:
        p.paragraph_format.space_before = Pt(10)
        run = p.add_run(text)
        set_font(run, Pt(12), bold=True)


def add_body(text):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = Pt(22)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    set_font(run)


def add_bold_body(text):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = Pt(22)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    set_font(run, bold=True)


def add_url(text):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = Pt(22)
    run = p.add_run(text)
    set_font(run, underline=True)


def add_table(headers, rows):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers), style='Table Grid')
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ''
        run = cell.paragraphs[0].add_run(h)
        set_font(run, bold=True)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = table.rows[r + 1].cells[c]
            cell.text = ''
            run = cell.paragraphs[0].add_run(val)
            set_font(run)
    doc.add_paragraph()


# ===================== 封面 =====================
add_title('教师作业批改助手')
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('产品说明书与操作指南')
set_font(run, Pt(14))

doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('创作人：马健驰    学号：20235963')
set_font(run)
doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run(f'版本：v1.1    更新日期：{datetime.date.today().strftime("%Y年%m月%d日")}')
set_font(run)
doc.add_page_break()

# ===================== 一、产品概述 =====================
add_title('一、产品概述', 1)
add_body('教师作业批改助手是一款基于大语言模型（LLM）的智能作业批改系统。教师只需输入学生作业内容、选择学科，AI 即可自动完成评分、逐题批改、优点分析、改进建议及教师寄语，大幅提升批改效率。')
add_body('本产品支持 11 门学科、3 家 AI 服务商（DeepSeek、Kimi 月之暗面、SiliconFlow 硅基流动），可手动输入文字或上传文件（.txt / .docx / .pdf / .md），并提供严格、标准、宽松三种批改尺度。')

# ===================== 二、公开访问链接 =====================
add_title('二、公开访问链接', 1)
add_body('以下链接可直接在浏览器中打开，无需安装任何软件：')

add_bold_body('[ AI 作业批改助手 - 主页面 ]')
add_url('https://source-commons-prefers-convergence.trycloudflare.com')
add_body('功能：输入/上传作业 - 选择学科和模型 - AI 批改 - 查看/导出结果')

add_bold_body('[ 产品说明书 - 在线版 ]')
add_url('https://source-commons-prefers-convergence.trycloudflare.com/manual')
add_body('功能：查看完整的产品说明和常见问题解答')

add_body('注意：以上链接为 Cloudflare Tunnel 动态地址，服务端保持运行期间有效。如链接失效，请联系管理员获取最新链接。推荐使用 Chrome 或 Edge 浏览器访问。')

# ===================== 三、快速上手 =====================
add_title('三、快速上手（5 步完成批改）', 1)

add_title('步骤 1：获取 API Key', 2)
add_body('在使用批改功能前，需要先获取 AI 服务商的 API Key（三选一即可）：')
add_table(
    ['服务商', '获取地址', '免费额度', '推荐模型'],
    [
        ['DeepSeek 官方', 'platform.deepseek.com', '新用户注册即送', 'DeepSeek-V3（综合最强）'],
        ['Moonshot Kimi', 'platform.moonshot.cn', '注册送 15 元体验金', 'Kimi K2.5（256K 上下文）'],
        ['SiliconFlow 硅基流动', 'cloud.siliconflow.cn', '免费模型可用', 'Qwen3-8B（免费）'],
    ]
)

add_title('步骤 2：打开公网链接', 2)
add_body('在浏览器中打开上方提供的公网链接，看到左侧配置面板 + 右侧作业输入区的界面。')

add_title('步骤 3：配置 API', 2)
add_body('(1) 在左侧"API 服务商"下拉框中选择你注册的服务商。注意：DeepSeek 的 Key 只能选"DeepSeek 官方"，不可与其他服务商混用。')
add_body('(2) 在"API Key"输入框中粘贴你的 Key。')
add_body('(3) 在"批改模型"中选择模型（推荐选项：DeepSeek-V3 或 Qwen3-8B）。')

add_title('步骤 4：输入作业并开始批改', 2)
add_body('(1) 在左侧"选择课程"中点击目标学科（如数学）。')
add_body('(2) 选择知识点/题型，以及批改标准（严格 / 标准 / 宽松）。')
add_body('(3) 在右侧输入框中粘贴作业内容，或切换到"上传文件"页签上传文档。')
add_body('(4) 点击右上角"开始批改"按钮。')
add_body('(5) 等待 10-30 秒，右侧将显示 AI 生成的结构化批改结果。')

add_title('步骤 5：查看与导出结果', 2)
add_body('批改结果包含：评分 - 总体评价 - 详细批改 - 优点 - 不足与建议 - 教师寄语。')
add_body('结果区下方提供三个操作按钮：复制结果（一键复制到剪贴板）、导出文本（下载为 .txt 文件）、打印（直接打印或保存为 PDF）。')

# ===================== 四、支持学科 =====================
add_title('四、支持学科一览', 1)
add_table(
    ['学科', '可选知识点'],
    [
        ['语文', '阅读理解、作文、古诗词鉴赏、文言文翻译、现代文阅读'],
        ['数学', '代数、几何、概率统计、函数、方程与不等式'],
        ['英语', '阅读理解、写作、完形填空、语法填空、翻译'],
        ['物理', '力学、电磁学、热学、光学、原子物理'],
        ['化学', '有机化学、无机化学、化学反应原理、实验题、化学方程式'],
        ['生物', '细胞生物学、遗传学、生态学、人体生理、进化论'],
        ['历史', '中国古代史、中国近代史、世界史、材料分析、论述题'],
        ['地理', '自然地理、人文地理、区域地理、地图分析、综合题'],
        ['政治', '经济生活、政治生活、文化生活、哲学、时政分析'],
        ['编程/信息', 'Python、C/C++、算法、数据结构、程序填空'],
        ['自定义课程', '自由输入学科名称和知识点'],
    ]
)

# ===================== 五、批改标准 =====================
add_title('五、批改标准说明', 1)
add_table(
    ['标准', '适用场景', '特点'],
    [
        ['严格', '考试评分', '不放过任何错误，扣分明确，适合模拟考'],
        ['标准（推荐）', '日常作业批改', '指出错误并给出改进建议，兼顾严谨与鼓励'],
        ['宽松', '鼓励性评价', '以肯定为主，温和指出不足，适合低年级学生'],
    ]
)

# ===================== 六、常见问题 (FAQ) =====================
add_title('六、常见问题 (FAQ)', 1)

faqs = [
    ('Q1: 提示"API 调用失败: Api key is invalid"？',
     '最常见的原因：服务商和 API Key 不匹配。DeepSeek 的 Key 只能选"DeepSeek 官方"，SiliconFlow 的 Key 只能选"SiliconFlow"，Kimi 的 Key 只能选"Moonshot (Kimi 月之暗面)"。请检查左侧是否选对了服务商。'),
    ('Q2: 提示"temperature: only 1 is allowed for this model"？',
     '某些推理模型（如 DeepSeek-R1、Kimi K2.5）只接受 temperature=1。切换到 DeepSeek-V3 或 Moonshot V1 等普通模型即可，系统已自动适配该参数。'),
    ('Q3: 左侧课程按钮点不了、科目全部消失了？',
     '通常是页面 JavaScript 加载异常。按 Ctrl+F5 强制刷新浏览器即可恢复。'),
    ('Q4: 批改结果中出现大量数学符号（如 $ \\frac \\Delta 等）？',
     'AI 有时会使用 LaTeX 数学格式。系统已优化提示词要求纯文本输出。如偶有出现，重新批改一次即可消除。'),
    ('Q5: 上传文件后在哪里查看文件内容？',
     '文件上传成功后，输入框下方会自动出现"文件内容预览"区域，展示解析出的完整文本，超出可滚动查看。'),
    ('Q6: 支持什么文件格式？',
     '支持 .txt、.docx、.pdf、.md 四种格式，推荐 .txt 和 .docx 效果最佳。'),
    ('Q7: API Key 安全吗？',
     '安全。Key 仅存浏览器本地（localStorage），直连 AI 服务商，不经过服务器记录或转发。'),
    ('Q8: 可以免费使用吗？',
     '可以。DeepSeek 新用户送额度、Kimi 注册送 15 元、SiliconFlow 有免费模型。'),
    ('Q9: 公网链接打不开怎么办？',
     '链接基于 Cloudflare Tunnel，服务端重启会导致地址变更。请联系管理员获取最新链接，推荐 Chrome / Edge 浏览器。'),
    ('Q10: Kimi 模型批改太慢？',
     '建议单题用 DeepSeek-V3 或 Moonshot V1-8K（更快），长篇作业再用 Kimi K2.5（256K 上下文优势）。系统已将输出限制为 2048 tokens。'),
    ('Q11: 批改结果准确吗？',
     'AI 批改仅供参考，最终评分请以教师判断为准。建议 AI 初筛后，教师重点复核 AI 标记的错误项和改进项。'),
]

for q, a in faqs:
    add_bold_body(q)
    add_body(a)

# ===================== 七、技术架构 =====================
add_title('七、技术架构', 1)
add_table(
    ['层级', '技术选型', '说明'],
    [
        ['前端', 'HTML + CSS + JavaScript（原生）', '无框架依赖，响应式设计'],
        ['后端', 'Python Flask', '请求转发、提示词工程、文件解析'],
        ['AI 引擎', 'OpenAI 兼容接口', '支持 DeepSeek / Kimi / SiliconFlow'],
        ['文件解析', 'python-docx + PyPDF2', 'Word 和 PDF 文本提取'],
        ['部署方式', '本地运行 + Cloudflare Tunnel', '公网 HTTPS 访问'],
    ]
)

# ===================== 八、测试样例 =====================
add_title('八、测试样例', 1)

add_title('样例 1：数学 - 代数（含错误）', 2)
add_body('[题目] 解方程 2x平方 - 5x - 3 = 0')
add_body('[学生解答] 使用求根公式，a=2, b=-5, c=-3，判别式 = 25+24 = 49，x1 = 3，x2 = 0.5')
add_body('提示：x2 应为 -0.5（符号错误），可测试 AI 能否发现。')

add_title('样例 2：地理 - 自然地理（含错误）', 2)
add_body('[题目] 简述影响气候的主要因素。')
add_body('[学生解答] 纬度、海陆位置、地形、洋流、大气环流。洋流：暖流增温增湿，寒流降温增湿。')
add_body('提示：寒流应为"降温减湿"而非"降温增湿"，可测试 AI 能否发现概念错误。')

add_title('样例 3：语文 - 作文片段', 2)
add_body('[题目] 以"我的老师"为题写一段话。')
add_body('[学生习作] 我的老师十分和蔼，她有一头乌黑的头发，脸上总挂着微笑。上课时她认真讲解每一个知识，下课后她耐心解答我们的问题。')
add_body('提示：可测试 AI 对写作的点评能力，如语句通顺度、描写生动性等。')

# ===================== 页脚 =====================
doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('--- END ---')
set_font(run)

# 保存
# 保存到项目目录（避免桌面文件被占用）
output = 'c:/Users/小灰灰/Desktop/src/homework-grader/教师作业批改助手_产品说明书.docx'
doc.save(output)
print('OK')
# 也尝试保存到桌面（新文件名避开旧文件）
import shutil
try:
    desktop_out = 'c:/Users/小灰灰/Desktop/教师作业批改助手_产品说明书_v2.docx'
    shutil.copy(output, desktop_out)
    print('Desktop:', desktop_out)
except Exception as e:
    print('Copy to desktop failed:', e)
