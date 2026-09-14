# -*- coding: utf-8 -*-
"""双语简历模板生成：中文 + 英文（同一数据模型，保证逐条同步）；顶部预留证件照占位框。"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_TAB_ALIGNMENT, WD_LINE_SPACING, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DIR = r"D:\内推\社招\chjiang\resume_template"
OUT = {"zh": DIR + r"\数字后端工程师简历模板.docx",
       "en": DIR + r"\数字后端工程师简历模板(英文版).docx"}

CN, CN_H, EN = "宋体", "黑体", "Arial"
BODY_PT = 10.5
DARK = RGBColor(0x22, 0x22, 0x22)
GRAY = RGBColor(0x55, 0x55, 0x55)
LINE_PT = 15

# ---------------- 通用标签 ----------------
LAB = {
 "zh": dict(obj="求职意向", skill="专业技能", work="工作经历", proj="项目经历",
            edu="教育经历", cert="荣誉证书 / 自我评价",
            comp="公司简介：", bg="项目背景：", duty="主要职责：", res="项目成果：",
            summary="自我评价：", photo="证件照", photo_tip="（单击此处插入照片）"),
 "en": dict(obj="Job Objective", skill="Technical Skills", work="Work Experience",
            proj="Key Projects", edu="Education", cert="Certifications & Summary",
            comp="Company: ", bg="Background: ", duty="Responsibilities:", res="Results:",
            summary="Summary: ", photo="PHOTO", photo_tip="(Insert photo here)"),
}

# ---------------- 顶部 ----------------
NAME = {"zh": "XX  X X", "en": "Firstname  Lastname"}
INFO = {"zh": "电话：XXXXX    邮箱：XX@XX.com    工作年限：3年    现居城市：XX市    求职状态：在职-考虑机会/离职-随时到岗",
        "en": "Phone: +86 XXXXX     Email: XX@XX.com     Experience: 3 years     Location: XX     Status: Open to opportunities"}

# ---------------- 求职意向 ----------------
OBJ_ROWS = {
 "zh": [("应聘职位：", "数字后端工程师", "期望行业：", "芯片/半导体"),
        ("期望城市：", "XX市", "期望薪资：", "XX-XXK/月")],
 "en": [("Target Role: ", "Digital Backend Engineer (Physical Design)", "Industry: ", "Semiconductor / IC"),
        ("Preferred City: ", "XX", "Expected Salary: ", "XX–XX K/month")],
}
OBJ_TAIL = {"zh": "到岗时间：XX周内到岗        工作性质：全职",
            "en": "Availability: Within XX weeks        Employment Type: Full-time"}

# ---------------- 专业技能 ----------------
SKILLS = [
 ("设计流程：", "熟悉数字芯片后端全流程（Netlist→GDSII），包括逻辑综合（Design Compiler/Genus）、Floorplan、电源网络规划、布局布线（ICC2/Innovus）、时钟树综合（CTS）、功能/时序 ECO；",
  "Design Flow: ", "Full digital backend flow from netlist to GDSII, including logic synthesis (Design Compiler/Genus), floorplanning, power planning, place & route (ICC2/Innovus), CTS, and functional/timing ECO;"),
 ("时序收敛：", "熟练使用 PrimeTime/Tempus 进行静态时序分析（STA），熟悉 MMMC 多模式多工艺角签核，能独立完成 setup/hold 违例分析与修复；",
  "Timing Closure: ", "Proficient in STA with PrimeTime/Tempus; experienced in MMMC multi-mode multi-corner signoff; independently analyze and fix setup/hold violations;"),
 ("时钟树：", "能独立完成 CTS 与时钟树优化，熟悉 useful skew、clock gating，可平衡 skew、功耗与时序；",
  "Clock Tree: ", "Independent CTS and optimization; familiar with useful skew and clock gating to balance skew, power and timing;"),
 ("物理验证：", "熟练使用 Calibre 完成 DRC/LVS/ERC/Antenna 检查与违例修复，保证 signoff 干净；",
  "Physical Verification: ", "Skilled in Calibre DRC/LVS/ERC/Antenna checks and violation fixing for clean signoff;"),
 ("低功耗设计：", "熟悉 UPF 低功耗流程、多电压域（MV）设计，了解隔离单元、电平转换单元、电源开关单元插入及功耗分析；",
  "Low Power: ", "Familiar with UPF flow and multi-voltage-domain design; isolation / level-shifter / power-switch cell insertion and power analysis;"),
 ("可靠性签核：", "了解 IR-Drop/EM 分析（Redhawk/Voltus），能协同修复电压降与电迁移违例；",
  "Reliability Signoff: ", "Working knowledge of IR-drop/EM analysis (Redhawk/Voltus) and collaborative violation fixing;"),
 ("验证与DFT：", "了解形式验证（Formality/Conformal）；能配合 DFT 完成扫描链插入、ATPG 及测试友好的物理实现；",
  "Verification & DFT: ", "Basic knowledge of formal verification (Formality/Conformal); cooperate with DFT on scan insertion, ATPG and DFT-friendly implementation;"),
 ("脚本能力：", "熟练使用 Tcl/Python/Perl/Shell 编写流程自动化与结果检查脚本，熟悉 Linux 环境与 Git 版本管理；",
  "Scripting: ", "Proficient in Tcl/Python/Perl/Shell for flow automation and result checking; comfortable with Linux and Git;"),
 ("工艺经验：", "具备 XXnm/XXnm 工艺节点后端实现经验，参与 X 颗芯片流片，其中 X 颗一次流片成功。【按实际删改】",
  "Process Experience: ", "Hands-on experience on XXnm/XXnm nodes; participated in X tape-outs with X first-pass success. [Edit as needed]"),
]

# ---------------- 工作经历 ----------------
T = {"zh": "XX年XX月 - XX年XX月", "en": "XX/20XX – XX/20XX"}
COMPANY1 = ("XX半导体（XX）有限公司", "XX Semiconductor Co., Ltd.")
TITLE1 = ("数字后端工程师  |  XX部门", "Digital Backend Engineer  |  XX Dept", "XX市", "XX")
COMPANY_INTRO = ("公司为专注于 XX（如 SoC/通信基带/AI 加速/MCU）芯片的 Fabless 设计企业，核心产品为 XX，团队约 XX 人。【一句话即可，可删】",
                 "A fabless IC company focusing on XX (SoC / baseband / AI accelerator / MCU) chips; core product XX; team of ~XX. [One sentence; deletable]")
WORK1 = [
 ("负责 XX 模块从 Netlist 到 GDSII 的后端物理实现，完成 Floorplan、电源网络、PnR、CTS 与 ECO，模块规模约 XX 万门、目标频率 XX MHz；",
  "Own backend implementation of the XX module from netlist to GDSII: floorplan, power network, PnR, CTS and ECO; scale ~XX M gates, target frequency XX MHz;"),
 ("建立并维护模块 MMMC 时序约束与 STA 环境，使用 PrimeTime 定位关键路径，通过 cell sizing、buffer 插入、布局/布线优化等手段完成 setup/hold 收敛，将 WNS 由 -XX ps 修复至 +XX ps；",
  "Built and maintained module MMMC constraints and STA environment; identified critical paths with PrimeTime; achieved setup/hold closure via cell sizing, buffer insertion and placement/route optimization, improving WNS from -XX ps to +XX ps;"),
 ("独立完成时钟树综合与优化，将 clock skew 控制在 XX ps 内、insertion delay 约 XX ps，并通过时钟门控降低时钟功耗 XX%；",
  "Performed CTS independently, keeping clock skew within XX ps and insertion delay ~XX ps; cut clock power by XX% via clock gating;"),
 ("使用 Calibre 完成 DRC/LVS/Antenna 检查，分析并修复 XX 条违例，最终实现零违例 signoff；",
  "Ran Calibre DRC/LVS/Antenna checks, analyzed and fixed XX violations, achieving zero-violation signoff;"),
 ("基于 UPF 完成多电压域实现及隔离/电平转换/电源开关单元插入，模块动态、静态功耗合计降低 XX%；",
  "Implemented multi-voltage-domain design with isolation / level-shifter / power-switch cells under UPF, reducing total power by XX%;"),
 ("使用 Tcl/Python 开发 XX 个自动化脚本，覆盖网表处理、数据准备与结果检查，后端交付效率提升 XX%；",
  "Developed XX Tcl/Python automation scripts for netlist processing, data prep and result checking, improving delivery efficiency by XX%;"),
 ("与前端、DFT、模拟及封装团队对接接口、约束与时序预算，参与设计评审，保障模块按计划节点交付。",
  "Coordinated with front-end, DFT, analog and package teams on interfaces, constraints and timing budgets; joined design reviews to ensure on-time delivery."),
]
HINT2 = {"zh": "【第二段工作经历模板，如只有一段请整体删除】",
         "en": "[Template for a second job — delete this block if not applicable]"}
COMPANY2 = ("XX科技有限公司", "XX Technology Co., Ltd.")
TITLE2 = ("助理/初级数字后端工程师", "Junior Digital Backend Engineer")
WORK2 = [
 ("协助完成 XX 模块布局布线、CTS 辅助优化及物理验证违例修复；",
  "Assisted in PnR, CTS optimization and physical-verification fixing of the XX module;"),
 ("负责流程数据准备、脚本维护与结果汇总，参与 X 次流片交付。",
  "Handled data preparation, script maintenance and result summary; participated in X tape-outs."),
]

# ---------------- 项目一 ----------------
P1_NAME = ("项目一：XXnm XX 芯片（SoC/通信基带）后端设计",
           "Project 1: XXnm XX Chip (SoC / Baseband) Backend Design")
P1_META = ("角色：核心模块后端负责人    规模：XX 万门 / 主频 XX MHz",
           "Role: Core-module backend lead     Scale: XX M gates / XX MHz", "工艺：XXnm", "Process: XXnm")
P1_BG = ("该芯片面向 XX 应用，集成 XX、XX 等模块，采用 XXnm 工艺，后端周期仅 XX 个月，面临高频、低功耗与布线拥塞等多重挑战。",
         "A chip for XX applications integrating XX and XX modules on XXnm; backend schedule only XX months, challenged by high frequency, low power and routing congestion.")
P1_DUTY = [
 ("负责 XX 核心模块后端全流程，制定 Floorplan 与电源网络方案，完成 PnR、CTS 与 ECO；",
  "Owned full-flow backend implementation of the core XX module: floorplan & power-plan strategy, PnR, CTS and ECO;"),
 ("主导模块时序收敛，分析多 corner 关键路径，分层修复 setup/hold 违例共 XX 条；",
  "Led timing closure; analyzed multi-corner critical paths and fixed XX setup/hold violations in stages;"),
 ("协同顶层完成模块集成与时序预算划分，解决接口处布线拥塞与物理验证问题；",
  "Cooperated with top-level on integration and timing budgeting; resolved interface congestion and PV issues;"),
 ("编写自动化检查脚本与交付 checklist，规范模块交付质量。",
  "Developed automation check scripts and delivery checklists to standardize quality."),
]
P1_RES = [
 ("模块频率达到 XX MHz，WNS/TNS 全部收敛，功耗较预算降低 XX%；",
  "Module reached XX MHz with full WNS/TNS closure; power XX% below budget;"),
 ("零 DRC/LVS 违例交付，支撑芯片一次流片成功且功能测试通过；",
  "Delivered with zero DRC/LVS violations, enabling first-pass tape-out success and passing functional tests;"),
 ("沉淀的流程与 checklist 被后续 X 个项目复用，节省约 XX 人日。",
  "Flow and checklists were reused by X following projects, saving ~XX person-days."),
]

# ---------------- 项目二 ----------------
P2_NAME = ("项目二：XX 模块时序收敛与低功耗/可靠性优化专项",
           "Project 2: Timing Closure, Low-Power & Reliability Optimization of XX Module")
P2_META = ("角色：专项优化负责人    规模：XX 万门",
           "Role: Optimization task-force lead     Scale: XX M gates", "工艺：XXnm", "Process: XXnm")
P2_BG = ("针对 XX 芯片 XX 模块在 XXnm 工艺下 WNS 违例严重、局部 IR-Drop 超标的问题，开展专项优化以保障流片节点。",
         "Dedicated optimization for the XX module on XXnm, which suffered severe negative WNS and local IR-drop violations, in order to secure the tape-out schedule.")
P2_DUTY = [
 ("重新梳理关键时序路径，结合 useful skew、cell sizing、布局与逻辑重组进行增量优化；",
  "Re-analyzed critical paths and performed incremental optimization via useful skew, cell sizing, placement and logic restructuring;"),
 ("优化时钟树结构与时钟门控，降低动态功耗；",
  "Optimized clock-tree structure and clock gating to reduce dynamic power;"),
 ("使用 Redhawk/Voltus 分析 IR-Drop/EM，通过补加电源 stripe、插入 decap 单元修复电压热点。",
  "Analyzed IR-drop/EM with Redhawk/Voltus; fixed hotspots by adding power stripes and decap cells."),
]
P2_RES = [
 ("WNS 由 -XX ps 提升至 +XX ps，hold 违例清零；",
  "WNS improved from -XX ps to +XX ps; all hold violations cleared;"),
 ("最差 IR-Drop 由 XX% 降至 XX% 以内，模块功耗降低 XX%；",
  "Worst IR-drop reduced from XX% to within XX%; module power down XX%;"),
 ("专项仅用 XX 周完成，保障芯片按时流片。",
  "Task finished in XX weeks, securing on-time tape-out."),
]

# ---------------- 教育 ----------------
EDU = ("XX大学（985/211/双一流）", "XX University (985/211/Double First-Class)")
EDU2 = ("XX专业  硕士研究生/本科", "XX, Master / Bachelor")
EDU_B = [
 ("GPA：XX/XX（专业前 XX%）；主修课程：数字集成电路设计、半导体物理、Verilog HDL、低功耗设计、CMOS 模拟电路等；【按实际填写】",
  "GPA: XX/XX (top XX%); key courses: Digital IC Design, Semiconductor Physics, Verilog HDL, Low-Power Design, CMOS Analog Circuits. [Edit as needed]"),
 ("在校荣誉：XX 奖学金、XX 竞赛 XX 奖。【无可删】",
  "Honors: XX Scholarship, XX Competition Award. [Delete if none]"),
]

# ---------------- 证书/自评 ----------------
CERT = [
 ("证书：XX EDA 工具认证、英语 CET-X（XX 分）、XX。【无可删】",
  "Certifications: XX EDA certification, CET-X (score XX), XX. [Delete if none]"),
]
SUMMARY = ("拥有 3 年数字后端设计经验，熟悉 Netlist→GDSII 全流程及主流 EDA 工具，具备 XXnm 工艺 X 颗芯片流片经验；在时序收敛、时钟树、物理验证与低功耗设计上可独立承担模块级交付；善用脚本提升流程效率，工作严谨、沟通协作顺畅，可快速上手新项目。",
           "3 years of digital backend design experience; familiar with the full netlist-to-GDSII flow and mainstream EDA tools; tape-out experience on XXnm across X chips; able to independently deliver module-level work in timing closure, clock tree, physical verification and low-power design; leverage scripting for efficiency; rigorous, collaborative and quick to ramp up on new projects.")


# ====================== 文档构建 ======================
def build(lang):
    L = LAB[lang]
    doc = Document()
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21.0), Cm(29.7)
    s.top_margin = s.bottom_margin = Cm(1.9)
    s.left_margin = s.right_margin = Cm(2.2)
    cw = s.page_width - s.left_margin - s.right_margin

    normal = doc.styles["Normal"]
    normal.font.name, normal.font.size = EN, Pt(BODY_PT)
    normal.font.color.rgb = DARK
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), CN)
    nf = normal.paragraph_format
    nf.space_before, nf.space_after = Pt(0), Pt(2)
    nf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    nf.line_spacing = Pt(LINE_PT)

    def style_run(run, cn=CN, en=EN, size=BODY_PT, bold=False, italic=False, color=DARK):
        run.font.name, run.font.size = en, Pt(size)
        run.font.bold, run.font.italic, run.font.color.rgb = bold, italic, color
        rpr = run._element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts"); rpr.append(rf)
        rf.set(qn("w:eastAsia"), cn)

    def mkpara(after=2, before=0, align=None):
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before, pf.space_after = Pt(before), Pt(after)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(LINE_PT)
        if align is not None: p.alignment = align
        return p

    def border_bottom(p, color="666666"):
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement("w:pBdr")
        b = OxmlElement("w:bottom")
        b.set(qn("w:val"), "single"); b.set(qn("w:sz"), "6")
        b.set(qn("w:space"), "3"); b.set(qn("w:color"), color)
        pBdr.append(b); pPr.append(pBdr)

    def title(text):
        p = mkpara(after=3, before=7)
        r = p.add_run(text); style_run(r, cn=CN_H, size=12.5, bold=True)
        border_bottom(p); return p

    def lr(left, right, bold=False, color=DARK, size=BODY_PT):
        p = mkpara()
        p.paragraph_format.tab_stops.add_tab_stop(cw, WD_TAB_ALIGNMENT.RIGHT)
        r1 = p.add_run(left); style_run(r1, bold=bold, size=size, color=color)
        r2 = p.add_run("\t" + right); style_run(r2, size=size, color=color)
        return p

    def bullet(pair, prefix=None):
        p = doc.add_paragraph(style="List Bullet")
        pf = p.paragraph_format
        pf.left_indent, pf.first_line_indent = Cm(0.62), Cm(-0.32)
        pf.space_after = Pt(1.5)
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(LINE_PT)
        if prefix:
            r0 = p.add_run(prefix); style_run(r0, bold=True)
        r = p.add_run(pair); style_run(r)

    def hint(text):
        p = mkpara(after=1)
        r = p.add_run(text); style_run(r, size=9, italic=True, color=GRAY)

    # ---------- 顶部：姓名信息(左) + 证件照(右) ----------
    photo_w = Cm(3.0)
    head = doc.add_table(rows=1, cols=2)
    head.autofit = False
    tblPr = head._tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr"); head._tbl.insert(0, tblPr)
    layout = OxmlElement("w:tblLayout"); layout.set(qn("w:type"), "fixed"); tblPr.append(layout)
    lc, rc = head.rows[0].cells
    lc.width, rc.width = cw - photo_w, photo_w
    # 行高（二寸照片比例约 3:4）
    trPr = head.rows[0]._tr.get_or_add_trPr()
    th = OxmlElement("w:trHeight"); th.set(qn("w:val"), str(int(Cm(3.8).twips))); th.set(qn("w:hRule"), "atLeast")
    trPr.append(th)
    rc.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    # 左：姓名 + 信息
    pn = lc.paragraphs[0]
    pn.paragraph_format.space_after = Pt(3)
    rn = pn.add_run(NAME[lang]); style_run(rn, cn=CN_H, size=21, bold=True)
    pi = lc.add_paragraph()
    pi.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pi.paragraph_format.line_spacing = Pt(LINE_PT)
    ri = pi.add_run(INFO[lang]); style_run(ri, size=10)
    # 右：证件照占位框（给单元格加细边框）
    pp = rc.paragraphs[0]; pp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pp.paragraph_format.space_after = Pt(0)
    r = pp.add_run(L["photo"]); style_run(r, size=10, bold=True, color=GRAY)
    pp2 = rc.add_paragraph(); pp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = pp2.add_run(L["photo_tip"]); style_run(r, size=8.5, color=GRAY)
    tcPr = rc._tc.get_or_add_tcPr()
    tcB = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        e = OxmlElement(f"w:{edge}")
        e.set(qn("w:val"), "single"); e.set(qn("w:sz"), "8")
        e.set(qn("w:space"), "0"); e.set(qn("w:color"), "999999")
        tcB.append(e)
    tcPr.append(tcB)

    # ---------- 求职意向 ----------
    title(L["obj"])
    t = doc.add_table(rows=2, cols=4)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for ri_, row in enumerate(t.rows):
        for ci_, cell in enumerate(row.cells):
            cp = cell.paragraphs[0]
            cp.paragraph_format.space_after = Pt(1)
            rr = cp.add_run(OBJ_ROWS[lang][ri_][ci_])
            style_run(rr, size=10, bold=(ci_ % 2 == 0))
    p = mkpara(after=2, before=1)
    r = p.add_run(OBJ_TAIL[lang]); style_run(r, size=10)

    # ---------- 专业技能 ----------
    title(L["skill"])
    for row in SKILLS:
        if lang == "zh":
            bullet(row[1], prefix=row[0])
        else:
            bullet(row[3], prefix=row[2])

    # ---------- 工作经历 ----------
    title(L["work"])
    lr(COMPANY1[0 if lang == "zh" else 1], T[lang], bold=True)
    lr(TITLE1[0 if lang == "zh" else 1], TITLE1[2 if lang == "zh" else 3], color=GRAY)
    p = mkpara(after=1)
    r0 = p.add_run(L["comp"]); style_run(r0, bold=True, size=10)
    r = p.add_run(COMPANY_INTRO[0 if lang == "zh" else 1]); style_run(r, size=10, color=GRAY)
    for w in WORK1: bullet(w[0 if lang == "zh" else 1])
    hint(HINT2[lang])
    lr(COMPANY2[0 if lang == "zh" else 1], T[lang], bold=True)
    lr(TITLE2[0 if lang == "zh" else 1], "", color=GRAY)
    for w in WORK2: bullet(w[0 if lang == "zh" else 1])

    # ---------- 项目 ----------
    def project(name, meta, bg, duty, res):
        lr(name[0 if lang == "zh" else 1], T[lang], bold=True)
        lr(meta[0 if lang == "zh" else 1], meta[2 if lang == "zh" else 3], color=GRAY)
        p = mkpara(after=1)
        r0 = p.add_run(L["bg"]); style_run(r0, bold=True, size=10)
        r = p.add_run(bg[0 if lang == "zh" else 1]); style_run(r, size=10)
        p = mkpara(after=1); r = p.add_run(L["duty"]); style_run(r, bold=True, size=10)
        for d in duty: bullet(d[0 if lang == "zh" else 1])
        p = mkpara(after=1); r = p.add_run(L["res"]); style_run(r, bold=True, size=10)
        for x in res: bullet(x[0 if lang == "zh" else 1])

    title(L["proj"])
    project(P1_NAME, P1_META, P1_BG, P1_DUTY, P1_RES)
    project(P2_NAME, P2_META, P2_BG, P2_DUTY, P2_RES)

    # ---------- 教育 ----------
    title(L["edu"])
    lr(EDU[0 if lang == "zh" else 1], T[lang], bold=True)
    lr(EDU2[0 if lang == "zh" else 1], "", color=GRAY)
    for b in EDU_B: bullet(b[0 if lang == "zh" else 1])

    # ---------- 证书/自评 ----------
    title(L["cert"])
    for c in CERT: bullet(c[0 if lang == "zh" else 1])
    p = mkpara(after=2)
    r0 = p.add_run(L["summary"]); style_run(r0, bold=True)
    r = p.add_run(SUMMARY[0 if lang == "zh" else 1]); style_run(r)

    doc.save(OUT[lang])
    print("saved:", OUT[lang])


for lg in ("zh", "en"):
    build(lg)
