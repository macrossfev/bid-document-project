#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
章节自动匹配引擎
根据章节类型，从资料模板库、人员库、业绩库中自动匹配推荐内容。
"""

from datetime import date, timedelta
from models import db, CompanyAttachment, Personnel, Performance


def _tag_overlap(tags_str_a, tags_str_b):
    """计算两个逗号分隔标签字符串的交集数量"""
    if not tags_str_a or not tags_str_b:
        return 0
    set_a = {t.strip().lower() for t in tags_str_a.split(',') if t.strip()}
    set_b = {t.strip().lower() for t in tags_str_b.split(',') if t.strip()}
    return len(set_a & set_b)


def _keyword_in_text(keywords, text):
    """检查关键词列表中有多少个出现在文本中"""
    if not text:
        return 0
    text_lower = text.lower()
    return sum(1 for kw in keywords if kw.lower() in text_lower)


def match_attachment(section_type, category):
    """
    根据章节类型匹配资料模板。

    Returns:
        list of dict: [{'id': int, 'name': str, 'category': str, 'score': float}, ...]
    """
    results = []

    # 先按 category 精确匹配
    exact_matches = CompanyAttachment.query.filter_by(category=category).all()
    for att in exact_matches:
        score = 80.0
        # 未过期的加分
        if att.expiry_date and att.expiry_date >= date.today():
            score += 10
        elif att.expiry_date and att.expiry_date < date.today():
            score -= 30  # 已过期降权
        # 有文件的加分
        if att.file_path:
            score += 10
        results.append({
            'id': att.id,
            'name': att.name,
            'category': att.category,
            'score': min(100, max(0, score)),
            'expired': bool(att.expiry_date and att.expiry_date < date.today()),
        })

    # 如果精确匹配不够，用 section_type 关键词做模糊匹配
    if len(results) < 3:
        all_atts = CompanyAttachment.query.filter(
            CompanyAttachment.category != category
        ).all()
        for att in all_atts:
            score = 0
            # 名称中包含 section_type
            if section_type and section_type in (att.name or ''):
                score += 40
            # tags 中包含 section_type
            if section_type and att.tags and section_type in att.tags:
                score += 30
            if score > 0:
                if att.file_path:
                    score += 10
                if att.expiry_date and att.expiry_date < date.today():
                    score -= 30
                results.append({
                    'id': att.id,
                    'name': att.name,
                    'category': att.category or '未分类',
                    'score': min(100, max(0, score)),
                    'expired': bool(att.expiry_date and att.expiry_date < date.today()),
                })

    # 按分数降序排列
    results.sort(key=lambda x: x['score'], reverse=True)
    return results[:5]


def match_personnel(project_tags=None):
    """
    为人员资料章节匹配推荐人员。

    Args:
        project_tags: 项目的行业标签字符串 (comma-separated)

    Returns:
        dict: {
            'project_leader': [{'id', 'name', 'title', 'score'}, ...],
            'tech_leader': [...],
            'members': [...],
        }
    """
    all_personnel = Personnel.query.all()
    project_tags_set = set()
    if project_tags:
        project_tags_set = {t.strip().lower() for t in project_tags.split(',') if t.strip()}

    scored = []
    for p in all_personnel:
        score = 0
        role_hint = ''

        # 职务匹配
        position = (p.position or '').lower()
        if '项目负责' in position or '项目经理' in position:
            role_hint = '项目负责人'
            score += 30
        elif '技术负责' in position or '技术总监' in position:
            role_hint = '技术负责人'
            score += 30

        # 职称加分
        title = (p.title or '').lower()
        if '高级' in title or '教授' in title or '正高' in title:
            score += 20
        elif '中级' in title or '工程师' in title:
            score += 10

        # 技能标签与项目标签交集
        if project_tags_set and p.skills:
            skills_set = {t.strip().lower() for t in p.skills.split(',') if t.strip()}
            overlap = len(project_tags_set & skills_set)
            score += overlap * 15

        # 有证书加分
        if p.certificates:
            score += min(len(p.certificates) * 5, 20)

        scored.append({
            'id': p.id,
            'name': p.name,
            'title': p.title or '',
            'position': p.position or '',
            'skills': p.skills or '',
            'role_hint': role_hint,
            'score': min(100, score),
        })

    scored.sort(key=lambda x: x['score'], reverse=True)

    # 分角色推荐
    result = {
        'project_leader': [],
        'tech_leader': [],
        'members': [],
    }

    for item in scored:
        if item['role_hint'] == '项目负责人':
            result['project_leader'].append(item)
        elif item['role_hint'] == '技术负责人':
            result['tech_leader'].append(item)
        else:
            result['members'].append(item)

    # 限制每类数量
    result['project_leader'] = result['project_leader'][:3]
    result['tech_leader'] = result['tech_leader'][:3]
    result['members'] = result['members'][:10]

    return result


def match_performance(project_tags=None, years=3):
    """
    为业绩章节匹配推荐业绩。

    Args:
        project_tags: 项目的行业标签字符串 (comma-separated)
        years: 近几年的业绩 (默认3年)

    Returns:
        list of dict: [{'id', 'project_name', 'client_name', 'score', ...}, ...]
    """
    cutoff_date = date.today() - timedelta(days=years * 365)
    all_perf = Performance.query.all()

    project_tags_set = set()
    if project_tags:
        project_tags_set = {t.strip().lower() for t in project_tags.split(',') if t.strip()}

    scored = []
    for perf in all_perf:
        score = 0

        # 时间范围内加分
        if perf.service_start and perf.service_start >= cutoff_date:
            score += 30
        elif perf.service_start and perf.service_start >= cutoff_date - timedelta(days=365):
            score += 15  # 稍早但仍有参考价值

        # 服务类型标签匹配
        if project_tags_set and perf.service_types:
            overlap = _tag_overlap(','.join(project_tags_set), perf.service_types)
            score += overlap * 20

        # 检测指标标签匹配
        if project_tags_set and perf.testing_params:
            overlap = _tag_overlap(','.join(project_tags_set), perf.testing_params)
            score += overlap * 10

        # 合同金额越大加分越多（适度）
        if perf.contract_amount and perf.contract_amount > 0:
            if perf.contract_amount >= 100:
                score += 15
            elif perf.contract_amount >= 50:
                score += 10
            else:
                score += 5

        # 有附件证明材料加分
        if perf.files:
            score += min(len(perf.files) * 5, 15)

        scored.append({
            'id': perf.id,
            'project_name': perf.project_name,
            'client_name': perf.client_name or '',
            'contract_amount': perf.contract_amount,
            'service_start': perf.service_start.isoformat() if perf.service_start else '',
            'service_end': perf.service_end.isoformat() if perf.service_end else '',
            'service_types': perf.service_types or '',
            'score': min(100, max(0, score)),
        })

    scored.sort(key=lambda x: x['score'], reverse=True)
    return scored[:10]


def match_requirements_to_sections(requirements_text, sections):
    """
    将完整要求文本与章节列表进行匹配。

    Args:
        requirements_text: 完整的招标要求文本
        sections: BidSection 对象列表

    Returns:
        dict: {section_id: matched_requirement_text}
    """
    if not requirements_text:
        return {}

    from tender_parser import SECTION_KEYWORD_MAP
    result = {}
    req_lines = requirements_text.split('\n')

    for section in sections:
        keywords = []
        sec_type = section.section_type or ''
        sec_name = section.section_name or ''

        # 从章节名提取中文关键词
        import re
        name_keywords = re.findall(r'[\u4e00-\u9fff]{2,}', sec_name)
        keywords.extend(name_keywords)

        # 从 SECTION_KEYWORD_MAP 获取关键词
        if sec_type in SECTION_KEYWORD_MAP:
            keywords.extend(SECTION_KEYWORD_MAP[sec_type]['keywords'])

        if not keywords:
            result[section.id] = ''
            continue

        matched_lines = []
        for line in req_lines:
            line = line.strip()
            if not line:
                continue
            for kw in keywords:
                if kw in line:
                    matched_lines.append(line)
                    break

        result[section.id] = '\n'.join(matched_lines[:10])

    return result


def _parse_requirement_keywords(requirement_text):
    """
    从要求文本中提取具体的要求关键词，用于增强匹配。
    例如: "需提供CMA证书" -> ['CMA', '证书']
          "项目负责人需具备高级职称" -> ['高级职称', '项目负责人']
    """
    if not requirement_text:
        return []

    import re
    keywords = []
    # 提取证书名称
    cert_patterns = [
        r'(CMA|CNAS|ISO\d*)',
        r'(资质认定|计量认证|检验检测)',
        r'(高级|中级|初级)(职称|工程师)',
        r'(营业执照|社保|纳税)',
        r'(合同|中标通知|验收报告)',
    ]
    for pat in cert_patterns:
        matches = re.findall(pat, requirement_text)
        for m in matches:
            if isinstance(m, tuple):
                keywords.append(''.join(m))
            else:
                keywords.append(m)

    # 提取中文关键词短语
    phrases = re.findall(r'需[提供具备有]([\u4e00-\u9fff]+)', requirement_text)
    keywords.extend(phrases)

    return keywords


def auto_match_section(section, project_tags=None):
    """
    对单个章节进行自动匹配。

    Args:
        section: BidSection 对象
        project_tags: 项目行业标签

    Returns:
        dict: {
            'attachment': {'id': int, 'name': str, 'score': float} or None,
            'personnel': {...} or None,
            'performances': [...] or None,
        }
    """
    result = {
        'attachment': None,
        'personnel': None,
        'performances': None,
    }

    sec_type = section.section_type or ''
    category = ''

    # 从 tender_parser 的映射中查找 category
    from tender_parser import SECTION_KEYWORD_MAP
    for stype, info in SECTION_KEYWORD_MAP.items():
        if stype == sec_type:
            category = info['category']
            break
    if not category:
        category = sec_type  # fallback

    # 如果章节有 requirement_text，提取额外关键词增强匹配
    req_keywords = []
    if hasattr(section, 'requirement_text') and section.requirement_text:
        req_keywords = _parse_requirement_keywords(section.requirement_text)

    # 匹配资料模板
    att_matches = match_attachment(sec_type, category)
    # 如果有要求关键词，对匹配结果重新评分
    if att_matches and req_keywords:
        for att_match in att_matches:
            bonus = _keyword_in_text(req_keywords, att_match.get('name', ''))
            att_match['score'] = min(100, att_match['score'] + bonus * 5)
        att_matches.sort(key=lambda x: x['score'], reverse=True)
    if att_matches:
        result['attachment'] = att_matches[0]  # 取最佳匹配

    # 人员类章节
    if sec_type in ('人员资料',):
        result['personnel'] = match_personnel(project_tags)

    # 业绩类章节
    if sec_type in ('业绩证明',):
        result['performances'] = match_performance(project_tags)

    return result


def auto_match_all_sections(bid_project):
    """
    对投标项目的所有章节进行自动匹配，并更新数据库。

    Args:
        bid_project: BidProject 对象

    Returns:
        dict: 匹配统计信息
    """
    from models import BidSection, BidPersonnel, BidPerformance

    project_tags = bid_project.industry_tags
    matched_count = 0

    for section in bid_project.sections:
        match = auto_match_section(section, project_tags)

        # 如果有资料模板匹配且当前未设置
        if match['attachment'] and not section.attachment_id:
            section.attachment_id = match['attachment']['id']
            section.match_score = match['attachment']['score']
            matched_count += 1

        # 如果是人员章节，自动关联推荐的人员
        if match['personnel']:
            existing_personnel_ids = {bp.personnel_id for bp in bid_project.bid_personnel}
            for role_key, role_name in [('project_leader', '项目负责人'),
                                         ('tech_leader', '技术负责人')]:
                for person in match['personnel'].get(role_key, [])[:1]:  # 每角色取第一个
                    if person['id'] not in existing_personnel_ids:
                        bp = BidPersonnel(
                            bid_project_id=bid_project.id,
                            personnel_id=person['id'],
                            role=role_name,
                        )
                        db.session.add(bp)
                        existing_personnel_ids.add(person['id'])

        # 如果是业绩章节，自动关联推荐的业绩
        if match['performances']:
            existing_perf_ids = {bp.performance_id for bp in bid_project.bid_performances}
            for perf in match['performances'][:3]:  # 取前3个
                if perf['id'] not in existing_perf_ids:
                    bperf = BidPerformance(
                        bid_project_id=bid_project.id,
                        performance_id=perf['id'],
                    )
                    db.session.add(bperf)
                    existing_perf_ids.add(perf['id'])

    db.session.commit()

    return {
        'matched_sections': matched_count,
        'total_sections': len(bid_project.sections),
    }
