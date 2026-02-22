#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""初始化系统数据 - 从投标文件和基本资料中提取"""

import os
import shutil
from datetime import date, datetime
from app import app, db
from models import (Personnel, Certificate, Performance, PerformanceFile,
                    CompanyAttachment, BidProject, BidPersonnel, BidPerformance)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MATERIAL_DIR = os.path.join(BASE_DIR, '基本资料')
CERT_DIR = os.path.join(MATERIAL_DIR, '职称证书扫描件(1)')
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')

COMPANY = '重庆水务集团股份有限公司水质检测分公司'


def copy_file(src_path, dest_subfolder):
    """复制文件到 uploads 目录，返回相对路径"""
    if not os.path.exists(src_path):
        return None
    dest_dir = os.path.join(UPLOAD_DIR, dest_subfolder)
    os.makedirs(dest_dir, exist_ok=True)
    filename = os.path.basename(src_path)
    dest_path = os.path.join(dest_dir, filename)
    if not os.path.exists(dest_path):
        shutil.copy2(src_path, dest_path)
    return os.path.join(dest_subfolder, filename)


def init_personnel():
    """录入人员数据"""
    if Personnel.query.count() > 0:
        print('人员数据已存在，跳过')
        return

    # 从投标文件和职称证书中提取的人员信息
    personnel_data = [
        {
            'name': '黄河笑', 'gender': '男',
            'title': '主要负责人', 'title_level': '',
            'position': '主要负责人/负责人',
            'social_security_unit': COMPANY,
            'skills': '企业管理,项目统筹',
            'resume': '重庆水务集团股份有限公司水质检测分公司主要负责人。',
        },
        {
            'name': '张逸林', 'gender': '',
            'title': '高级工程师', 'title_level': '高级',
            'specialty': '环境/检测',
            'position': '项目负责人',
            'social_security_unit': COMPANY,
            'skills': '水质检测,项目管理,质量控制,污水检测,水库水质检测',
            'resume': '高级工程师，担任项目负责人，全面负责项目的组织协调、质量控制、进度管理和对外联络工作。具备丰富的水质检测项目管理经验。',
            'cert_file': '张逸林高级工程师.pdf',
        },
        {
            'name': '李亚莹', 'gender': '女',
            'title': '', 'title_level': '',
            'position': '投标委托代理人',
            'social_security_unit': COMPANY,
            'skills': '招投标管理,项目协调',
            'resume': '担任投标委托代理人，负责投标活动的组织实施和对外联络。',
            'cert_file': '李亚莹.pdf',
        },
        {
            'name': '贾海舰', 'gender': '',
            'title': '高级工程师', 'title_level': '高级',
            'specialty': '环境检测',
            'position': '技术骨干',
            'social_security_unit': COMPANY,
            'skills': '水质检测,技术管理,数据审核',
            'resume': '高级工程师，技术骨干，具备丰富的水质检测技术经验。',
            'cert_file': '贾海舰高级工程师.pdf',
        },
        {
            'name': '黄抒', 'gender': '',
            'title': '高级工程师', 'title_level': '高级',
            'specialty': '环境检测',
            'position': '技术骨干',
            'social_security_unit': COMPANY,
            'skills': '水质检测,实验室分析',
            'resume': '高级工程师，技术骨干。',
            'cert_file': '黄抒高级工程师.pdf',
        },
        {
            'name': '习苏芸', 'gender': '',
            'title': '', 'title_level': '',
            'position': '检测人员',
            'social_security_unit': COMPANY,
            'skills': '水质检测',
            'cert_file': '习苏芸.pdf',
        },
        {
            'name': '冯琳', 'gender': '',
            'title': '', 'title_level': '',
            'position': '检测人员',
            'social_security_unit': COMPANY,
            'skills': '水质检测',
            'cert_file': '冯琳.pdf',
        },
        {
            'name': '印成', 'gender': '',
            'title': '', 'title_level': '',
            'position': '检测人员',
            'social_security_unit': COMPANY,
            'skills': '水质检测',
            'cert_file': '印成.jpg',
        },
        {
            'name': '周燕', 'gender': '女',
            'title': '', 'title_level': '',
            'position': '检测人员',
            'social_security_unit': COMPANY,
            'skills': '水质检测',
            'cert_file': '周燕.pdf',
        },
        {
            'name': '周金元', 'gender': '',
            'title': '', 'title_level': '',
            'position': '检测人员',
            'social_security_unit': COMPANY,
            'skills': '水质检测',
            'cert_file': '周金元.pdf',
        },
        {
            'name': '姚思佳', 'gender': '',
            'title': '', 'title_level': '',
            'position': '检测人员',
            'social_security_unit': COMPANY,
            'skills': '水质检测',
            'cert_file': '姚思佳.pdf',
        },
        {
            'name': '张继蓉', 'gender': '女',
            'title': '', 'title_level': '',
            'position': '检测人员',
            'social_security_unit': COMPANY,
            'skills': '水质检测',
            'cert_file': '张继蓉.pdf',
        },
        {
            'name': '朱仁庆', 'gender': '',
            'title': '', 'title_level': '',
            'position': '检测人员',
            'social_security_unit': COMPANY,
            'skills': '水质检测',
            'cert_file': '朱仁庆.pdf',
        },
        {
            'name': '李显芳', 'gender': '',
            'title': '', 'title_level': '',
            'position': '检测人员',
            'social_security_unit': COMPANY,
            'skills': '水质检测',
            'cert_file': '李显芳.pdf',
        },
        {
            'name': '杨涛', 'gender': '男',
            'title': '', 'title_level': '',
            'position': '检测人员',
            'social_security_unit': COMPANY,
            'skills': '水质检测',
            'cert_file': '杨涛.jpg',
        },
        {
            'name': '柏雪', 'gender': '女',
            'title': '', 'title_level': '',
            'position': '检测人员',
            'social_security_unit': COMPANY,
            'skills': '水质检测',
            'cert_file': '柏雪.pdf',
        },
        {
            'name': '江珊珊', 'gender': '女',
            'title': '', 'title_level': '',
            'position': '检测人员',
            'social_security_unit': COMPANY,
            'skills': '水质检测',
            'cert_file': '江珊珊.jpg',
        },
        {
            'name': '熊伟丽', 'gender': '女',
            'title': '', 'title_level': '',
            'position': '检测人员',
            'social_security_unit': COMPANY,
            'skills': '水质检测',
            'cert_file': '熊伟丽.pdf',
        },
        {
            'name': '胡晓玲', 'gender': '女',
            'title': '', 'title_level': '',
            'position': '检测人员',
            'social_security_unit': COMPANY,
            'skills': '水质检测',
            'cert_file': '胡晓玲.jpg',
        },
        {
            'name': '钟声', 'gender': '',
            'title': '', 'title_level': '',
            'position': '检测人员',
            'social_security_unit': COMPANY,
            'skills': '水质检测',
            'cert_file': '钟声.jpg',
        },
        {
            'name': '陈洋', 'gender': '',
            'title': '', 'title_level': '',
            'position': '检测人员',
            'social_security_unit': COMPANY,
            'skills': '水质检测',
            'cert_file': '陈洋.jpg',
        },
        {
            'name': '陈茜', 'gender': '女',
            'title': '工程师', 'title_level': '中级',
            'position': '检测人员',
            'social_security_unit': COMPANY,
            'skills': '水质检测',
            'cert_file': '陈茜中级职称证书.pdf',
        },
        {
            'name': '雷湘', 'gender': '',
            'title': '', 'title_level': '',
            'position': '检测人员',
            'social_security_unit': COMPANY,
            'skills': '水质检测',
            'cert_file': '雷湘.pdf',
        },
    ]

    for data in personnel_data:
        cert_file = data.pop('cert_file', None)
        specialty = data.pop('specialty', None)
        if specialty:
            data['specialty'] = specialty

        p = Personnel(**data)
        db.session.add(p)
        db.session.flush()  # get p.id

        # 添加职称证书
        if cert_file:
            src = os.path.join(CERT_DIR, cert_file)
            rel_path = copy_file(src, 'personnel')
            cert_type = '职称证' if '高级' in (data.get('title') or '') or '中级' in (data.get('title_level') or '') else '职称证'
            cert = Certificate(
                personnel_id=p.id,
                cert_type=cert_type,
                cert_name=f"{data['name']}职称证书",
                file_path=rel_path,
            )
            db.session.add(cert)

    db.session.commit()
    print(f'已录入 {len(personnel_data)} 名人员')


def init_performance():
    """录入业绩数据"""
    if Performance.query.count() > 0:
        print('业绩数据已存在，跳过')
        return

    performances = [
        {
            'project_name': '重庆水务集团所属污水厂出水水质抽检项目',
            'client_name': '重庆水务环境控股集团有限公司',
            'service_start': date(2025, 1, 1),
            'service_end': date(2025, 12, 31),
            'service_types': '污水检测,第三方抽检',
            'testing_params': 'COD,BOD5,TN,TP,氨氮,SS',
            'description': '对重庆水务集团所属污水厂出水进行水质第三方检测，检测指标包括COD、BOD5、TN、TP、氨氮、SS等。',
        },
        {
            'project_name': '重庆市城镇污水处理厂监督性监测项目',
            'client_name': '重庆市生态环境监测中心',
            'service_start': date(2024, 1, 1),
            'service_end': date(2025, 12, 31),
            'service_types': '污水检测,监督性监测',
            'testing_params': 'COD,BOD5,TN,TP,氨氮,SS',
            'description': '对重庆市城镇污水处理厂出水进行监督性检测，为环境监管提供数据支撑。',
        },
        {
            'project_name': '重庆市饮用水水源地水质监测项目',
            'client_name': '重庆市水利局',
            'service_start': date(2024, 1, 1),
            'service_end': date(2025, 12, 31),
            'service_types': '水源地检测,饮用水检测',
            'testing_params': '浑浊度,色度,pH,高锰酸盐指数,氨氮,总磷,总氮,叶绿素a,藻类',
            'description': '对重庆市饮用水水源地（水库）进行水质检测，保障供水水源安全。',
        },
        {
            'project_name': '重庆市农村生活污水处理设施运行监测项目',
            'client_name': '重庆市生态环境局',
            'service_start': date(2025, 1, 1),
            'service_end': date(2025, 12, 31),
            'service_types': '农村污水检测,运行监测',
            'testing_params': 'COD,BOD5,TN,TP,氨氮,SS',
            'description': '对重庆市农村生活污水处理设施出水进行运行监测，评估设施运行效果。',
        },
    ]

    for data in performances:
        perf = Performance(**data)
        db.session.add(perf)

    db.session.commit()
    print(f'已录入 {len(performances)} 条业绩')


def init_attachments():
    """录入公司附件"""
    if CompanyAttachment.query.count() > 0:
        print('附件数据已存在，跳过')
        return

    attachments = [
        {
            'name': '营业执照',
            'category': '营业执照',
            'src_file': '营业执照（新）.pdf',
            'notes': '统一社会信用代码：91500108MAE5G7MX3H，类型：股份有限公司分公司（上市、国有控股），负责人：黄河笑',
        },
        {
            'name': 'CMA资质认定证书',
            'category': 'CMA证书',
            'src_file': '资质认定证书（2021年）.pdf',
            'issue_date': date(2021, 12, 1),
            'expiry_date': date(2027, 11, 30),
            'notes': '证书编号：210013061568，检测能力覆盖水和废水、生活饮用水、地表水、地下水等水质检测',
        },
        {
            'name': 'CMA资质认定证书附表',
            'category': 'CMA证书',
            'src_file': '检验检测机构资质认定证书附表 (2025.1).pdf',
            'issue_date': date(2025, 1, 1),
            'notes': '检验检测机构资质认定证书能力附表，详列所有认定的检测参数和方法',
        },
        {
            'name': '房屋租赁合同',
            'category': '其他',
            'src_file': '房屋租赁合同.pdf',
            'notes': '实验室及办公场所租赁合同',
        },
        {
            'name': '授权委托书',
            'category': '授权委托书',
            'src_file': '授权委托书扫描件.pdf',
            'notes': '法定代表人授权委托书扫描件',
        },
        {
            'name': '负责人身份证明',
            'category': '其他',
            'src_file': '王总身份证扫描件.pdf',
            'notes': '公司负责人身份证扫描件',
        },
        {
            'name': '信用查询记录',
            'category': '其他',
            'src_file': '重庆水务集团股份有限公司水质检测分公司 (信用查询记录).pdf',
            'notes': '企业信用查询记录，证明无重大违法失信记录',
        },
        {
            'name': '社会保险参保证明',
            'category': '其他',
            'src_file': '重庆市社会保险参保证明（单位）—参保人员明细.pdf',
            'notes': '单位社保参保人员明细',
        },
        {
            'name': '总公司改革方案通知',
            'category': '授权委托书',
            'src_file': '水务集团《突出主责主业、增强核心功能整合优化改革方案》的通知.pdf',
            'notes': '水务集团关于突出主责主业、增强核心功能整合优化改革方案的通知，证明分公司设立背景',
        },
        {
            'name': '基本资格条件承诺函',
            'category': '其他',
            'src_file': '基本资格条件承诺函.doc',
            'notes': '投标基本资格条件承诺函',
        },
    ]

    for data in attachments:
        src_file = data.pop('src_file')
        src_path = os.path.join(MATERIAL_DIR, src_file)
        rel_path = copy_file(src_path, 'attachments')
        data['file_path'] = rel_path
        att = CompanyAttachment(**data)
        db.session.add(att)

    db.session.commit()
    print(f'已录入 {len(attachments)} 个附件')


def init_bid_project():
    """创建投标项目并关联人员和业绩"""
    if BidProject.query.count() > 0:
        print('投标项目已存在，跳过')
        return

    bid = BidProject(
        project_name='重庆水务环境控股集团有限公司2026年所属厂站生产指标专项抽检项目（包2）',
        bidder_name='重庆水务环境控股集团有限公司',
        agent_name='重庆水务集团公用工程咨询有限公司',
        service_period='2026年2月至2026年12月',
        status='in_progress',
        notes='包2：环投集团污水厂846座出水水质抽检 + 供水水库53座水质检测。检测指标：污水厂（COD、BOD5、TN、TP、氨氮、SS），水库（浑浊度、色度、嗅和味、肉眼可见物、pH、高锰酸盐指数、氨氮、总磷、总氮、叶绿素a、藻类）。',
    )
    db.session.add(bid)
    db.session.flush()

    # 关联项目负责人 - 张逸林
    pm = Personnel.query.filter_by(name='张逸林').first()
    if pm:
        db.session.add(BidPersonnel(bid_project_id=bid.id, personnel_id=pm.id, role='项目负责人'))

    # 关联委托代理人 - 李亚莹
    agent = Personnel.query.filter_by(name='李亚莹').first()
    if agent:
        db.session.add(BidPersonnel(bid_project_id=bid.id, personnel_id=agent.id, role='委托代理人'))

    # 关联技术骨干
    for name in ['贾海舰', '黄抒']:
        p = Personnel.query.filter_by(name=name).first()
        if p:
            db.session.add(BidPersonnel(bid_project_id=bid.id, personnel_id=p.id, role='技术骨干'))

    # 关联所有业绩
    for perf in Performance.query.all():
        db.session.add(BidPerformance(bid_project_id=bid.id, performance_id=perf.id))

    db.session.commit()
    print(f'已创建投标项目并关联人员和业绩')


if __name__ == '__main__':
    with app.app_context():
        print('=== 开始初始化系统数据 ===')
        init_personnel()
        init_performance()
        init_attachments()
        init_bid_project()
        print('=== 数据初始化完成 ===')
        print(f'人员: {Personnel.query.count()} 人')
        print(f'证书: {Certificate.query.count()} 个')
        print(f'业绩: {Performance.query.count()} 条')
        print(f'附件: {CompanyAttachment.query.count()} 个')
        print(f'投标项目: {BidProject.query.count()} 个')
