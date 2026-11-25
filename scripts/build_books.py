# -*- coding: utf-8 -*-
"""学生错题诊断流水线 · 双份 Word 生成器（家长版诊断报告 + 学生大白话对照本）

用法：
  python scripts/build_books.py --data <结构化JSON> [--out <输出目录>] [--work <临时目录>] [--only both|parent|student]

设计：模板在本文件，卷子数据在 JSON。换一套卷子只替换 --data，不改代码。
输出前跑 v2.0 核验字段门禁（v1.0 数据自动跳过）；输出后做三色 run 级自检与隐私扫描。
任一失败退出码为 1，未 PASS 不得交付。
"""
import os
import re
import sys
import json
import glob
import argparse
from datetime import date
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import docx_kit as k

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(HERE)
_DATA_CANDIDATES = (os.path.join(SKILL, 'templates', '*.json'),
                    os.path.join(SKILL, 'references', 'sample-*.json'))
_data_paths = [p for pattern in _DATA_CANDIDATES for p in sorted(glob.glob(pattern))]
DEFAULT_DATA = _data_paths[0] if _data_paths else None
DEFAULT_OUT = os.environ.get(
    'CUOTI_OUTPUT_DIR',
    os.path.join(os.getcwd(), 'output', date.today().strftime('%Y-%m-%d')))


def short_no(no, student=False):
    m = re.search(r'第\s*(\d+)\s*题', no)
    m_cls = re.search(r'第\s*(\d+)\s*类', no)
    if m:
        n = m.group(1)
    elif m_cls:
        n = m_cls.group(1) + '类'
    else:
        n = no[:6]
    m2 = re.search(r'第\s*\((\d+)\)\s*问', no)
    if m2:
        n += '(%s)' % m2.group(1)
    return n


def brief(text, limit=14):
    return text.split('（')[0].strip()[:limit]


def four_brief(four):
    parts = re.findall(r'[：:]\s*([^｜|]+)', four)
    clean = [re.split(r'[（(]', p)[0].strip()[:4] for p in parts[:4]]
    return '｜'.join(clean) if clean else four[:22]


def cause_brief(cause):
    letters = re.findall(r'(?:^|[＋+＋\s])([A-E])\s', cause)
    if not letters:
        letters = re.findall(r'([A-E])\s', cause)
    return '＋'.join(dict.fromkeys(letters)) if letters else cause[:8]


def verdict_rows(items):
    rows = []
    for it in items:
        rows.append([short_no(it['no']), brief(it['mine']), brief(it['right']),
                     it['score'], four_brief(it['four']), cause_brief(it['cause'])])
    return rows


def overview_rows(items):
    rows = []
    for it in items:
        rows.append([short_no(it['no'], True), brief(it['mine']), brief(it['right']),
                     it['score'], it['oneline'][:34] + '…'])
    return rows


AUDIT_STATUS = ('VALID', 'SUSPECT', 'INVALID')
CONSENSUS_SET = ('CONFIRMED', 'CONTESTED', 'PENDING')
CONFIDENCE_SET = ('high', 'medium', 'pending')


def _qno(text):
    m = re.search(r'第\s*(\d+)\s*题', str(text or ''))
    return m.group(1) if m else str(text or '').strip()


def _pct(text):
    m = re.search(r'(\d+(?:\.\d+)?)\s*%', str(text or ''))
    return float(m.group(1)) if m else None


def audit_checks(d):
    """v2.0 核验字段自检。返回 (state, summary, errors)。

    state = SKIPPED（v1.0 数据，无核验字段）/ PASS / FAIL。
    字段缺失时整块跳过，保证黄金样例与旧卷数据不受影响。
    """
    qa = d.get('question_audit') or []
    av = d.get('answer_verification') or []
    acc = d.get('accuracy_summary') or {}
    if not (qa or av or acc):
        return 'SKIPPED', 'v1.0 data', []

    errors = []
    counts = {'VALID': 0, 'SUSPECT': 0, 'INVALID': 0}
    for q in qa:
        st = str(q.get('status', '')).strip().upper()
        if st not in AUDIT_STATUS:
            errors.append('题面状态非法 %r @ %s（须为 VALID/SUSPECT/INVALID）' % (q.get('status'), q.get('no')))
        else:
            counts[st] += 1

    high = contested = n_contested = 0
    for a in av:
        c = str(a.get('consensus', '')).strip().upper()
        f = str(a.get('confidence', '')).strip().lower()
        if c not in CONSENSUS_SET:
            errors.append('收敛状态非法 %r @ %s' % (a.get('consensus'), a.get('no')))
        if f not in CONFIDENCE_SET:
            errors.append('置信度非法 %r @ %s' % (a.get('confidence'), a.get('no')))
        if c == 'CONFIRMED' and f == 'high':
            high += 1
        if c == 'CONTESTED':
            n_contested += 1
        if c in ('CONTESTED', 'PENDING') or f == 'pending':
            contested += 1
        if c == 'CONTESTED' and f == 'high':
            errors.append('CONTESTED 却标 high @ %s' % a.get('no'))

    invalid_nos = {_qno(q.get('no')) for q in qa
                   if str(q.get('status', '')).strip().upper() == 'INVALID'}
    item_nos = {_qno(it.get('no')) for it in d.get('parent_items', [])}
    item_nos |= {_qno(it.get('no')) for it in d.get('student_items', [])}
    mixed = sorted(invalid_nos & item_nos)
    if mixed:
        errors.append('INVALID 题混入逐题错题列表：%s' % '、'.join(mixed))

    if acc:
        try:
            total = int(acc.get('total', 0))
            counted = int(acc.get('counted', 0))
            correct = int(acc.get('correct', 0))
            ex = acc.get('excluded') or {}
            exsum = sum(int(v) for v in ex.values())
            if counted != total - exsum:
                errors.append('counted(%d) != total(%d) - excluded(%d)' % (counted, total, exsum))
            if counted > total or correct > counted:
                errors.append('counted > total 或 correct > counted')
            if 'invalid' in ex and int(ex.get('invalid', 0)) != counts['INVALID']:
                errors.append('excluded.invalid(%s) 与 INVALID 题数(%d) 不一致' % (ex.get('invalid'), counts['INVALID']))
            if 'suspect' in ex and int(ex.get('suspect', 0)) != counts['SUSPECT']:
                errors.append('excluded.suspect(%s) 与 SUSPECT 题数(%d) 不一致' % (ex.get('suspect'), counts['SUSPECT']))
            if 'contested' in ex and int(ex.get('contested', 0)) < n_contested:
                errors.append('excluded.contested 少于 CONTESTED 题数，有分歧题未从分母剔除')
        except (TypeError, ValueError) as exc:
            errors.append('accuracy_summary 数值错误：%s' % exc)
        for key in ('accuracy', 'coverage', 'convergence'):
            if not acc.get(key):
                errors.append('accuracy_summary 缺 %s（三项须同时给出）' % key)
        gate = str(acc.get('gate', '')).strip().upper()
        if gate not in ('PASS', 'HOLD'):
            errors.append('gate 非法 %r（须为 PASS/HOLD）' % acc.get('gate'))
        cv, cg = _pct(acc.get('coverage')), _pct(acc.get('convergence'))
        if gate == 'PASS':
            if cv is not None and cv < 80:
                errors.append('gate=PASS 但覆盖率 %.0f%% < 80%%' % cv)
            if cg is not None and cg < 90:
                errors.append('gate=PASS 但答案收敛率 %.0f%% < 90%%' % cg)
            if '待裁定' in str(acc.get('accuracy', '')):
                errors.append('gate=PASS 但正确率栏写了「待裁定」')
        if gate == 'HOLD' and '待裁定' not in str(acc.get('accuracy', '')):
            errors.append('gate=HOLD 时正确率栏须写「待裁定，样本不足不予计算」')

    summary = ('valid=%d suspect=%d invalid=%d | confirmed/high=%d contested=%d'
               ' | accuracy=%s coverage=%s convergence=%s'
               % (counts['VALID'], counts['SUSPECT'], counts['INVALID'], high, contested,
                  acc.get('accuracy', '-'), acc.get('coverage', '-'), acc.get('convergence', '-')))
    return ('FAIL' if errors else 'PASS'), summary, errors


def render_audit_parent(doc, d):
    """家长版 v2.0 章节：题面审查 + 答案交叉论证 + 正确率口径。字段缺失则整章跳过。"""
    qa = d.get('question_audit') or []
    av = d.get('answer_verification') or []
    acc = d.get('accuracy_summary') or {}
    dq = d.get('doubts_question') or []
    da = d.get('doubts_answer') or []
    if not (qa or av or acc or dq or da):
        return

    k.h1(doc, '三·补、题面有效性与答案交叉论证（v2.0）')
    k.para(doc, '本节回答两个问题：卷子上的题本身能不能做；判对错所依据的标准答案是否经多轮独立论证。学生版只呈现结论提示，不呈现论证过程。',
           space_after=5)

    if qa:
        k.h3(doc, '一、题面有效性审查（六维 × 五视角）')
        rows = [[short_no(q.get('no', '')), str(q.get('status', '')).strip().upper(),
                 q.get('dimensions', ''), q.get('evidence', ''), q.get('suggestion', '')]
                for q in qa]
        k.table(doc, ['题号', '状态', '六维审查', '判定证据', '处置建议'], rows,
                [1.8, 2.0, 3.6, 4.4, 4.0])
        counts = {'VALID': 0, 'SUSPECT': 0, 'INVALID': 0}
        for q in qa:
            st = str(q.get('status', '')).strip().upper()
            if st in counts:
                counts[st] += 1
            if st != 'VALID':
                label = '【题目有误】' if st == 'INVALID' else '【题面存疑】'
                p = k.para(doc, '', space_after=3)
                k.set_run(p.add_run(label), k.TEA, 10.5, True)
                k.set_run(p.add_run('%s：%s　建议：%s'
                                    % (q.get('no', ''), q.get('evidence', ''), q.get('suggestion', ''))),
                          k.TEA, 10.5)
        k.para(doc, '题面三态计数：VALID %d 题 / SUSPECT %d 题 / INVALID %d 题。INVALID 题不判对错、不计入正确率、不入错题库、不出变式。'
               % (counts['VALID'], counts['SUSPECT'], counts['INVALID']), k.GREY, 9.5, space_after=6)

    if av:
        k.h3(doc, '二、标准答案交叉论证（六路独立求解＋红队反证）')
        rows = [[short_no(a.get('no', '')), str(a.get('consensus', '')).strip().upper(),
                 str(a.get('confidence', '')).strip().lower(),
                 '通过' if a.get('machine_verified') else '无／未通过',
                 a.get('routes', ''), a.get('verdict_note', '')]
                for a in av]
        k.table(doc, ['题号', '收敛', '置信度', '机器验证', '六路执行', '裁决说明'], rows,
                [1.6, 2.0, 1.8, 1.8, 4.4, 4.2])
        for a in av:
            if str(a.get('consensus', '')).strip().upper() == 'CONTESTED':
                k.para(doc, '· 【待裁定】%s：%s（分歧点：%s）'
                       % (a.get('no', ''), a.get('verdict_note', ''), a.get('divergence', '')),
                       k.TEA, 10, space_after=3, indent=0.3)
        k.para(doc, '收敛口径：≥ 4 路独立求解一致且红队未推翻才标 CONFIRMED；出现分歧一律标 CONTESTED 交人工裁定，禁止多数票定案。',
               k.GREY, 9.5, space_after=6)

    if acc:
        k.h3(doc, '三、正确率口径（三项同时给出）')
        k.table(doc, ['总题数', '计入统计', '判对题数', '正确率', '覆盖率', '答案收敛率', '输出门槛'],
                [[acc.get('total', ''), acc.get('counted', ''), acc.get('correct', ''),
                  acc.get('accuracy', ''), acc.get('coverage', ''), acc.get('convergence', ''),
                  acc.get('gate', '')]],
                [1.8, 2.0, 1.8, 2.2, 2.2, 2.4, 3.0])
        ex = acc.get('excluded') or {}
        k.para(doc, '剔除项：题目出错 %s 题、题面存疑 %s 题、答案分歧 %s 题、识别存疑 %s 题。正确率与覆盖率必须同时看；覆盖率低说明样本不足，不能只报正确率。'
               % (ex.get('invalid', 0), ex.get('suspect', 0), ex.get('contested', 0), ex.get('recognition', 0)),
               k.BODY, 10, space_after=4)
        if str(acc.get('gate', '')).strip().upper() == 'HOLD':
            k.para(doc, '【输出门槛未达标】本次不对外给出正确率结论，只交付判定表与存疑清单；待用户／老师裁定后整体重算。',
                   k.TEA, 10.5, True, space_after=6)

    if dq or da:
        k.h3(doc, '四、新增存疑清单')
        if dq:
            k.para(doc, '题面存疑：', k.NOTE, 10.5, True, space_after=2)
            for s in dq:
                k.para(doc, '· ' + str(s), k.BODY, 10, space_after=3, indent=0.4)
        if da:
            k.para(doc, '答案源存疑：', k.NOTE, 10.5, True, space_after=2)
            for s in da:
                k.para(doc, '· ' + str(s), k.BODY, 10, space_after=3, indent=0.4)


def _student_notices(d):
    notices = [str(x) for x in (d.get('student_question_notice') or [])]
    qa = d.get('question_audit') or []
    bad = [q for q in qa if str(q.get('status', '')).strip().upper() != 'VALID']
    if bad and not notices:
        for q in bad:
            n = short_no(q.get('no', ''), True)
            notices.append('第 %s 题，卷子自己可能印错了或者条件没给够，这道题先不算你错，等老师确认再说。' % n)
    return notices


def render_student_notice(doc, d):
    notices = _student_notices(d)
    if not notices:
        return
    k.h2(doc, '先说一件要紧的事')
    k.para(doc, '下面这几道题，卷子自己可能出错了。这几道先不算你错，等老师确认了再说：',
           k.TEA, 11, True, space_after=4)
    for s in notices:
        k.para(doc, '· ' + s, k.TEA, 10.5, space_after=3, indent=0.3)
    k.para(doc, '其余的题照常看，别因为这几道打乱节奏。', k.GREY, 9.5, space_after=8)


def render_parent(d):
    doc = Document()
    k.setup(doc)
    m = d['meta']
    k.para(doc, m['exam'], k.HEAD, 18, True, WD_ALIGN_PARAGRAPH.CENTER, 4)
    k.para(doc, '错题三色提取与深度诊断报告（家长版）', k.HEAD, 17, True, WD_ALIGN_PARAGRAPH.CENTER, 6)
    k.para(doc, '识别口径：印刷体＝试卷原题｜手写＝学生作答（报告中以深蓝标示）｜红色＝教师批改与正确答案',
           k.GREY, 9.5, False, WD_ALIGN_PARAGRAPH.CENTER, 4)
    k.para(doc, '生成日期：%s　｜　依据：%s' % (m['date'], m['source']),
           k.GREY, 9.5, False, WD_ALIGN_PARAGRAPH.CENTER, 12)
    k.para(doc, m['privacy_note'], k.GREY, 9, False, WD_ALIGN_PARAGRAPH.CENTER, 10)

    k.h1(doc, '一、三色说明与判定口径')
    k.para(doc, '本报告按三色提取口径分层转写，同一道题的三层信息用颜色区分，便于一眼定位错处：', space_after=5)
    k.kv(doc, '■ 黑色（#000000）＝', '印刷原题，照原文转写，不改写。', k.INK)
    k.kv(doc, '■ 深蓝（#003399）＝', '该生手写作答与卷面演算，含涂改标注。', k.STU)
    k.kv(doc, '■ 红色（#C00000）＝', '教师批改、红字订正与官方参考答案。', k.TEA)
    k.kv(doc, '■ 橙色（#E67E22）＝', '本报告的分析标注与提示。', k.NOTE)
    k.para(doc, '判定四维：答案（对／错）、方法（正确／错误／无法判断）、条件（完整／遗漏／无法判断）、表达（完整／有问题／无法判断），逐题给出，缺一不可。错因统一分五类：A 概念缺失／混淆、B 审题失误、C 方法·步骤、D 计算·操作、E 表达·规范。', space_after=6)

    k.h1(doc, '二、卷面得分概览')
    k.table(doc, m.get('score_headers', ['板块', '题号', '满分', '卷面得分', '失分', '失分题']),
            d['score_rows'], [3.0, 2.2, 1.8, 2.2, 1.8, 4.5])
    k.para(doc, m.get('score_note', '说明：得分取自答题卡各题左侧红色批改分数相加，为卷面记录值，仅供复盘参考。'), k.GREY, 9.5, space_after=6)

    k.h1(doc, '三、错题判定总表（先过一眼）')
    k.table(doc, m.get('verdict_headers', ['题号', '我的答案', '正确答案', '得分', '四维判定（简）', '错因']),
            verdict_rows(d['parent_items']), [1.8, 3.4, 3.6, 1.5, 4.2, 2.0])
    k.para(doc, m.get('reminder_line', '按技能口径，「答案对但过程可疑」单列提醒、不写入错题库；本报告仍纳入分析，因为它暴露了知识点或作答习惯上的隐患。'), k.GREY, 9.5, space_after=6)

    render_audit_parent(doc, d)

    k.h1(doc, '四、逐题深度诊断')
    for it in d['parent_items']:
        k.h2(doc, '%s　%s' % (it['no'], it['topic']))
        k.kv(doc, '【印刷题干】', it['stem'], k.INK)
        k.kv(doc, '【我的作答】', it['mine'], k.STU)
        k.kv(doc, '【正确答案／教师批改】', it['right'], k.TEA)
        k.kv(doc, '【卷面得分】', it['score'], k.TEA)
        p = k.para(doc, '', space_after=4)
        k.set_run(p.add_run('【错在哪】'), k.NOTE, 10.5, True)
        k.set_run(p.add_run(it['where']), k.TEA, 10.5, True)
        k.kv(doc, '【四维判定】', it['four'])
        k.kv(doc, '【错因归类】', it['cause'], k.NOTE)
        k.h3(doc, '正确思路与步骤')
        for i, s in enumerate(it['steps'], 1):
            k.para(doc, '%d. %s' % (i, s), k.BODY, 10.5, space_after=3, indent=0.4)
        k.h3(doc, '举一反三（变式自测）')
        for q, _ in it['variants']:
            k.para(doc, q, k.BODY, 10.5, space_after=2, indent=0.4)
            k.para(doc, '→ 答案见文末「变式答案（家长保管）」，自测时请先独立完成。',
                   k.GREY, 9, space_after=4, indent=0.8)
        k.kv(doc, '【对应知识点】', it['points'], k.HEAD)
        k.para(doc, '', space_after=6)

    k.h1(doc, '五、答案正确但需留意（不入错题库）')
    for label, text in d['reminders']:
        k.kv(doc, label + '：', text, k.STU)
    k.para(doc, '说明：以上按流水线口径列为「答案对但过程／痕迹可疑」的提醒项，不与错题混同。', k.GREY, 9.5, space_after=6)

    k.h1(doc, '六、知识点清单与薄弱点分布')
    k.table(doc, m.get('knowledge_headers', ['知识模块', '对应题目', '掌握情况', '薄弱点定位']),
            d['knowledge_rows'], [4.2, 3.0, 2.8, 6.5])
    k.para(doc, m.get('structural_line', '结构性结论：本次 %s 分失分集中在模型识别与过程书写，而非计算能力；提升重点是「模型识别」与「规范表达」，不是加大刷题量。' % m['lost']),
           k.BODY, 10.5, space_after=6)

    k.h1(doc, '七、提升方向与方法')
    k.table(doc, ['优先级', '主攻模块', '对应错题', '具体动作（每天 20–30 分钟）', '验收标准'],
            d['improve_rows'], [1.5, 3.0, 2.4, 7.0, 4.5])
    k.para(doc, '三色复盘法：黑笔抄题干、蓝笔写自己的答案、红笔写订正与关键依据；每题只在红笔旁写一句「错在第几步」。重做前先遮住红笔部分。', k.BODY, 10.5, space_after=4)
    k.para(doc, '错题本使用纪律：原题做对 ≠ 掌握。必须在变式自测全对后才标「已阻断」；看过答案或非限时完成的正确只记「辅助正确」。', k.BODY, 10.5, space_after=6)

    k.h1(doc, '八、复习排期（艾宾浩斯 1/2/4/7/15/30 天）')
    k.table(doc, ['日期', '节点', '任务'], k.schedule_rows(d['parent_review_tasks'], m['date']),
            [3.0, 1.8, 11.5])
    k.para(doc, '已阻断考点按 3／15／30 天三轮防复发；中途再错则回到错因归因，重新排期。', k.BODY, 10.5, space_after=6)

    k.h1(doc, '九、存疑清单（需人工复核）')
    k.h3(doc, '识别存疑')
    for s in d['doubts_recognition']:
        k.para(doc, '· ' + s, k.BODY, 10, space_after=3, indent=0.4)
    k.h3(doc, '判定存疑')
    for s in d['doubts_judgement']:
        k.para(doc, '· ' + s, k.BODY, 10, space_after=3, indent=0.4)
    if d.get('doubts_question'):
        k.h3(doc, '题面存疑')
        for s in d['doubts_question']:
            k.para(doc, '· ' + s, k.BODY, 10, space_after=3, indent=0.4)
    if d.get('doubts_answer'):
        k.h3(doc, '答案源存疑')
        for s in d['doubts_answer']:
            k.para(doc, '· ' + s, k.BODY, 10, space_after=3, indent=0.4)
    k.para(doc, '', space_after=6)

    k.h1(doc, '十、下一步确认事项')
    k.para(doc, '1. 请先核对第三节判定总表与第九节存疑清单；确认无误后回复「确认入库」，再写入错题库并生成带自动验证的变式训练卷。', space_after=4)
    k.para(doc, '2. 入库时还需补充两个必填字段：本次是否限时完成、是否偷看过答案。这两项会影响掌握判定，本次未默认填写。', space_after=4)
    k.para(doc, '3. 变式答案放在文末，请家长单独保管；学生自测时不要翻看。配套学生版对照本另行交付。', space_after=4)
    k.para(doc, '4. 若需要重拍，建议俯拍、光线均匀、字迹完整，尤其是红色与黑色笔迹叠加区域。', space_after=6)

    k.h1(doc, '附：变式答案（家长保管，学生自测时请勿翻看）')
    for it in d['parent_items']:
        k.h3(doc, it['no'].split('（')[0])
        for q, a in it['variants']:
            k.para(doc, q, k.GREY, 9.5, space_after=1)
            k.para(doc, a, k.TEA, 9.5, True, space_after=4, indent=0.4)
    k.para(doc, '—— 本报告依据照片与官方参考答案生成，未对原题与学生作答做内容改写。',
           k.GREY, 9, False, WD_ALIGN_PARAGRAPH.CENTER, 2)
    return doc


def render_student(d):
    doc = Document()
    k.setup(doc)
    m = d['meta']
    k.para(doc, '给学生看的错题本', k.HEAD, 19, True, WD_ALIGN_PARAGRAPH.CENTER, 4)
    k.para(doc, '%s（大白话版）' % m['exam'], k.HEAD, 16, True, WD_ALIGN_PARAGRAPH.CENTER, 4)
    k.para(doc, '看法很简单：先看清楚字 → 再判对错 → 最后把「我当时怎么想」和「应该怎么想」摆在一起比',
           k.GREY, 10, False, WD_ALIGN_PARAGRAPH.CENTER, 6)
    k.para(doc, '%s　｜　%s 道失分题　｜　卷面 %s / %s' % (m['date'], m['wrong_count'], m['score'], m['full_score']),
           k.GREY, 9.5, False, WD_ALIGN_PARAGRAPH.CENTER, 10)

    k.h2(doc, '先说三句话，看完再往下翻')
    k.para(doc, '1. 黑字是卷子上印的题，蓝字是我写的，红字是老师改的和正确答案。看题先看这三行。', space_after=3)
    k.para(doc, '2. 每道题都有一张「两条路」的表：左边是我当时怎么想的，右边是正确该怎么想的。红字那一格，就是我走岔的那一步。', space_after=3)
    k.para(doc, '3. 变式题先自己做完，答案在最后「对答案」那一节，别提前翻。', space_after=3)
    k.para(doc, '这本子里没有名字、班级、学号，只有题和怎么把题做对。', k.GREY, 9.5, space_after=8)

    render_student_notice(doc, d)

    k.h1(doc, '一、一眼看懂：我这次栽在哪')
    k.table(doc, m.get('overview_headers', ['题号', '我写的', '正确答案', '得分', '一句话说清']),
            overview_rows(d['student_items']), [1.6, 3.2, 3.2, 1.8, 6.2])
    k.para(doc, m.get('student_overview_line', '看这张表就够知道：不是计算不行，是「路子选错」和「过程写不清」丢的分最多。'),
           k.TEA, 10.5, True, space_after=8)

    k.h1(doc, '二、我的错都错在哪一类')
    k.table(doc, m.get('cause_headers', ['错因类型', '大白话是啥意思', '这次哪几题', '下次怎么防']),
            d['cause_rows'], [2.6, 5.0, 3.0, 5.4])
    k.para(doc, m.get('student_cause_line', '一句话总结：E 类「写不清」是我最容易白丢分的地方，而且最好补——不用多聪明，只要按格式写。'),
           k.TEA, 10.5, True, space_after=8)

    k.h1(doc, '三、逐道题，一道一道摆开看')
    for it in d['student_items']:
        k.h2(doc, '%s　%s　（%s）' % (it['no'], it['tag'], it['score']))
        k.kv(doc, '【这题在考啥】', it['what'], k.NOTE)
        k.kv(doc, '【题目】', it['stem'], k.INK)
        k.kv(doc, '【我写的】', it['mine'], k.STU)
        k.kv(doc, '【正确答案 / 老师改的】', it['right'], k.TEA)
        k.big(doc, '【一句话说清我错在哪】', it['oneline'])
        k.h3(doc, '两条路摆在一起看')
        k.contrast_table(doc, it['rows'])
        k.kv(doc, '【记一句口诀】', it['kou'], k.NOTE)
        k.kv(doc, '【错因归类】', it['cause'], k.NOTE)
        k.h3(doc, '换个数，再来两道')
        for q, _ in it['variants']:
            k.para(doc, q, k.BODY, 10, space_after=2, indent=0.4)
        k.para(doc, '（答案在最后「对答案」那一节，先自己做完。）', k.GREY, 9, space_after=4, indent=0.4)
        k.kv(doc, '【课本知识点】', it['point'], k.HEAD)
        k.kv(doc, '【怎么练最快】', it['drill'], k.NOTE)
        k.para(doc, '', space_after=6)

    k.h1(doc, '四、这些题到底考课本哪一块')
    k.table(doc, ['知识块', '牵连到的题', '我缺的是啥', '补它的最短动作'],
            d['student_knowledge_rows'], [4.0, 2.8, 4.2, 5.0])

    k.h1(doc, '五、每天 20 分钟，怎么练最快')
    k.table(doc, ['时间', '练什么', '怎么做', '练到什么程度算过'],
            d['drill_rows'], [1.8, 3.6, 5.6, 5.0])
    k.para(doc, m.get('student_drill_line', '顺序别换：先补「写清楚」，它最便宜、见效最快；再补「找相似」，它一次能救回好几道题的分。'),
           k.TEA, 10.5, True, space_after=6)
    k.h2(doc, '复习打卡表（做完打勾）')
    k.table(doc, ['日期', '第几天', '任务', '完成打勾'],
            [[r[0], r[1].replace('D+', '第 ') + ' 天', r[2], '□']
             for r in k.schedule_rows(d['student_review_tasks'], m['date'])],
            [3.0, 2.0, 9.5, 2.0])
    k.para(doc, '提醒：原题做对不算会。变式题全对、而且能讲给别人听，才算这题真的过去了。', k.GREY, 9.5, space_after=8)

    k.h1(doc, '六、对答案（做完再看）')
    for it in d['student_items']:
        k.h3(doc, it['no'])
        for q, a in it['variants']:
            k.para(doc, q, k.GREY, 9.5, space_after=1, indent=0.2)
            k.para(doc, '答：' + a, k.TEA, 9.5, True, space_after=4, indent=0.5)

    k.h1(doc, '七、几句实在话')
    for s in d['closing_lines']:
        k.para(doc, '· ' + s, space_after=3)
    k.h2(doc, '需要你核一下的地方')
    for s in d['student_doubts']:
        k.para(doc, '· ' + s, k.GREY, 9.5, space_after=2, indent=0.3)
    k.para(doc, '—— 题目和我的答案都按卷子原样抄，没改过；身份信息一律没写进本子。',
           k.GREY, 9, False, WD_ALIGN_PARAGRAPH.CENTER, 2)
    return doc


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--data', default=DEFAULT_DATA)
    ap.add_argument('--out', default=DEFAULT_OUT)
    ap.add_argument('--work', default=None)
    ap.add_argument('--only', default='both', choices=['both', 'parent', 'student'])
    ap.add_argument('--forbidden', default='', help='额外禁写的身份信息，逗号分隔')
    args = ap.parse_args()
    if not args.data or not os.path.isfile(args.data):
        print('ERROR: data json not found:', args.data)
        return 1
    work = args.work or os.path.join(args.out, '_work')
    with open(args.data, encoding='utf-8') as f:
        d = json.load(f)

    audit_state, audit_summary, audit_errors = audit_checks(d)
    if audit_state == 'SKIPPED':
        print('audit checks: SKIPPED (v1.0 data)')
    else:
        print('audit checks: %s | %s' % (audit_state, audit_summary))
    if audit_errors:
        for err in audit_errors:
            print('  audit FAIL:', err)
        print('RESULT: FAIL | files: 0 (audit gate blocked, no delivery)')
        return 1

    stem = re.sub(r'[\\/:*?"<>|\s·]+', '', d['meta']['exam']) + '_' + d['meta']['date']
    jobs = []
    if args.only in ('both', 'parent'):
        jobs.append(('parent', render_parent(d), stem + '_家长版三色诊断报告.docx'))
    if args.only in ('both', 'student'):
        jobs.append(('student', render_student(d), stem + '_学生大白话对照本.docx'))

    failed = False
    for tag, doc, filename in jobs:
        local, target = k.save_pair(doc, filename, work, args.out)
        counts, missing = k.color_audit(local)
        hits = k.privacy_scan(local, [x for x in args.forbidden.split(',') if x])
        print('[%s] %s' % (tag, target))
        print('  colors:', counts)
        print('  tri-color missing:', missing or 'none', '| privacy hits:', hits or 'none')
        if missing or hits:
            failed = True
    print('RESULT:', 'FAIL' if failed else 'PASS', '| files:', len(jobs))
    return 1 if failed else 0


if __name__ == '__main__':
    sys.exit(main())
