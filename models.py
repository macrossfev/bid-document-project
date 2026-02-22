from datetime import datetime, date
from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


class Personnel(db.Model):
    """人员库"""
    __tablename__ = 'personnel'

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(50), nullable=False)
    gender = db.Column(db.String(10))
    phone = db.Column(db.String(20))

    title = db.Column(db.String(100))              # 职称
    position = db.Column(db.String(100))            # 职务
    skills = db.Column(db.Text)                     # 能力标签, comma-separated

    # Keep old columns so SQLite doesn't break on existing data
    id_card = db.Column(db.String(18))
    title_level = db.Column(db.String(20))
    specialty = db.Column(db.String(100))
    education = db.Column(db.String(50))
    university = db.Column(db.String(100))
    work_years = db.Column(db.Integer)
    social_security_unit = db.Column(db.String(200))
    resume = db.Column(db.Text)

    created_at = db.Column(db.DateTime, default=datetime.now)
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    certificates = db.relationship(
        'Certificate', backref='personnel', lazy=True, cascade='all, delete-orphan'
    )

    def __repr__(self):
        return f'<Personnel {self.name}>'


class Certificate(db.Model):
    """人员证书"""
    __tablename__ = 'certificate'

    id = db.Column(db.Integer, primary_key=True)
    personnel_id = db.Column(db.Integer, db.ForeignKey('personnel.id'), nullable=False)
    cert_type = db.Column(db.String(50))    # 职称证/注册证/上岗证/学历证/身份证
    cert_name = db.Column(db.String(200))
    cert_number = db.Column(db.String(100))
    issue_date = db.Column(db.Date)
    expiry_date = db.Column(db.Date)
    file_path = db.Column(db.String(500))

    created_at = db.Column(db.DateTime, default=datetime.now)

    def __repr__(self):
        return f'<Certificate {self.cert_name}>'


class Performance(db.Model):
    """业绩库"""
    __tablename__ = 'performance'

    id = db.Column(db.Integer, primary_key=True)
    project_name = db.Column(db.String(300), nullable=False)
    client_name = db.Column(db.String(200))          # 委托单位
    contract_amount = db.Column(db.Float)             # 万元

    service_start = db.Column(db.Date)
    service_end = db.Column(db.Date)
    service_types = db.Column(db.Text)                # 服务类型标签
    testing_params = db.Column(db.Text)               # 检测指标标签
    description = db.Column(db.Text)

    created_at = db.Column(db.DateTime, default=datetime.now)
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    files = db.relationship(
        'PerformanceFile', backref='performance', lazy=True, cascade='all, delete-orphan'
    )

    def __repr__(self):
        return f'<Performance {self.project_name}>'


class PerformanceFile(db.Model):
    """业绩附件"""
    __tablename__ = 'performance_file'

    id = db.Column(db.Integer, primary_key=True)
    performance_id = db.Column(db.Integer, db.ForeignKey('performance.id'), nullable=False)
    file_type = db.Column(db.String(50))    # 合同/发票/中标通知书/验收报告
    file_name = db.Column(db.String(300))
    file_path = db.Column(db.String(500))

    created_at = db.Column(db.DateTime, default=datetime.now)

    def __repr__(self):
        return f'<PerformanceFile {self.file_name}>'


class CompanyAttachment(db.Model):
    """资料模板库"""
    __tablename__ = 'company_attachment'

    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), nullable=False)
    category = db.Column(db.String(50))     # 信誉承诺书/投标函模板/法定代表人证明/授权委托书/技术方案/营业执照/CMA证书/业绩证明/人员资料/报价表/其他
    file_path = db.Column(db.String(500))
    issue_date = db.Column(db.Date)
    expiry_date = db.Column(db.Date)
    version = db.Column(db.String(50))
    tags = db.Column(db.Text)
    notes = db.Column(db.Text)

    created_at = db.Column(db.DateTime, default=datetime.now)
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    def __repr__(self):
        return f'<CompanyAttachment {self.name}>'


class BidProject(db.Model):
    """投标项目"""
    __tablename__ = 'bid_project'

    id = db.Column(db.Integer, primary_key=True)
    project_name = db.Column(db.String(300), nullable=False)
    bidder_name = db.Column(db.String(200))          # 招标人
    agent_name = db.Column(db.String(200))           # 代理机构
    max_price = db.Column(db.Float)
    deadline = db.Column(db.DateTime)
    service_period = db.Column(db.String(100))

    project_type = db.Column(db.String(100))         # 项目类型: 环境检测/工程监理/...
    industry_tags = db.Column(db.Text)               # 行业标签, comma-separated

    tender_file_path = db.Column(db.String(500))          # 保留兼容，旧的完整招标文件
    tender_requirements = db.Column(db.Text)             # 招标要求部分文本
    tender_format = db.Column(db.Text)                   # 投标文件格式部分文本
    format_file_path = db.Column(db.String(500))         # 格式模板文件路径（用户拆分后上传）
    requirements_file_path = db.Column(db.String(500))   # 要求文档文件路径（用户拆分后上传）
    output_file_path = db.Column(db.String(500))
    status = db.Column(db.String(20), default='draft')  # draft/parsing/in_progress/completed
    notes = db.Column(db.Text)

    created_at = db.Column(db.DateTime, default=datetime.now)
    updated_at = db.Column(db.DateTime, default=datetime.now, onupdate=datetime.now)

    bid_personnel = db.relationship(
        'BidPersonnel', backref='bid_project', lazy=True, cascade='all, delete-orphan'
    )
    bid_performances = db.relationship(
        'BidPerformance', backref='bid_project', lazy=True, cascade='all, delete-orphan'
    )
    sections = db.relationship(
        'BidSection', backref='bid_project', lazy=True, cascade='all, delete-orphan'
    )

    def __repr__(self):
        return f'<BidProject {self.project_name}>'


class BidPersonnel(db.Model):
    """投标项目选用人员"""
    __tablename__ = 'bid_personnel'

    id = db.Column(db.Integer, primary_key=True)
    bid_project_id = db.Column(db.Integer, db.ForeignKey('bid_project.id'), nullable=False)
    personnel_id = db.Column(db.Integer, db.ForeignKey('personnel.id'), nullable=False)
    role = db.Column(db.String(50))         # 项目负责人/技术负责人/采样人员/检测人员

    personnel = db.relationship('Personnel', lazy=True)

    def __repr__(self):
        return f'<BidPersonnel project={self.bid_project_id} person={self.personnel_id}>'


class BidPerformance(db.Model):
    """投标项目选用业绩"""
    __tablename__ = 'bid_performance'

    id = db.Column(db.Integer, primary_key=True)
    bid_project_id = db.Column(db.Integer, db.ForeignKey('bid_project.id'), nullable=False)
    performance_id = db.Column(db.Integer, db.ForeignKey('performance.id'), nullable=False)

    performance = db.relationship('Performance', lazy=True)

    def __repr__(self):
        return f'<BidPerformance project={self.bid_project_id} perf={self.performance_id}>'


class BidSection(db.Model):
    """投标文件章节"""
    __tablename__ = 'bid_section'

    id = db.Column(db.Integer, primary_key=True)
    bid_project_id = db.Column(db.Integer, db.ForeignKey('bid_project.id'), nullable=False)
    section_key = db.Column(db.String(50), nullable=False)     # e.g. "section_1"
    section_name = db.Column(db.String(200), nullable=False)   # display name
    section_order = db.Column(db.Integer, default=0)           # 排序序号
    section_type = db.Column(db.String(50))                    # 标准类别: 投标函/报价表/营业执照/...
    source = db.Column(db.String(20), default='manual')        # manual/parsed (来源)
    original_requirement = db.Column(db.Text)                  # 招标文件中的原始要求文本
    requirement_text = db.Column(db.Text)                      # 从招标要求部分匹配到的具体要求
    attachment_id = db.Column(db.Integer, db.ForeignKey('company_attachment.id'), nullable=True)
    custom_content = db.Column(db.Text, nullable=True)
    match_score = db.Column(db.Float, default=0)               # 自动匹配置信度 0-100
    status = db.Column(db.String(20), default='pending')       # pending / done
    format_heading_text = db.Column(db.String(500))            # 格式模板中的章节标题原文（用于定位）
    format_para_index = db.Column(db.Integer)                  # 格式模板中章节标题的段落索引

    attachment = db.relationship('CompanyAttachment', lazy=True)

    def __repr__(self):
        return f'<BidSection {self.section_key}>'
