import os
import json
from datetime import datetime, date, timedelta
from functools import wraps
from flask import Flask, render_template, request, redirect, url_for, flash, send_file, session, jsonify
from werkzeug.utils import secure_filename
from config import Config
from models import (db, Personnel, Certificate, Performance, PerformanceFile,
                    CompanyAttachment, BidProject, BidPersonnel, BidPerformance, BidSection)
from tender_parser import (parse_tender_file, parse_tender_file_dual, get_default_sections,
                           parse_requirements_file, match_requirements_to_sections)
from section_matcher import auto_match_all_sections, match_attachment, match_personnel, match_performance

app = Flask(__name__)
app.config.from_object(Config)
db.init_app(app)

ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'pdf', 'doc', 'docx', 'xls', 'xlsx'}
TENDER_EXTENSIONS = {'docx', 'pdf'}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def save_upload(file, subfolder):
    if file and file.filename and allowed_file(file.filename):
        # 先提取原始扩展名（secure_filename 会丢弃中文字符导致扩展名丢失）
        original_ext = ''
        if '.' in file.filename:
            original_ext = file.filename.rsplit('.', 1)[1].lower()
        filename = secure_filename(file.filename)
        if not filename or '.' not in filename:
            # secure_filename 去掉了中文后可能只剩 'docx' 这样的无点字符串
            filename = f"upload.{original_ext}" if original_ext else 'upload'
        timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
        filename = f"{timestamp}_{filename}"
        folder = os.path.join(app.config['UPLOAD_FOLDER'], subfolder)
        os.makedirs(folder, exist_ok=True)
        filepath = os.path.join(folder, filename)
        file.save(filepath)
        return os.path.join(subfolder, filename)
    return None


def parse_date(val):
    if val:
        try:
            return datetime.strptime(val, '%Y-%m-%d').date()
        except ValueError:
            return None
    return None


def parse_datetime(val):
    if val:
        try:
            return datetime.strptime(val, '%Y-%m-%dT%H:%M')
        except ValueError:
            return None
    return None


def parse_float(val):
    if val:
        try:
            return float(val)
        except ValueError:
            return None
    return None


def parse_int(val):
    if val:
        try:
            return int(val)
        except ValueError:
            return None
    return None


with app.app_context():
    db.create_all()
    # Migrate: add new columns to existing tables if missing
    _migrate_columns = [
        ('bid_project', 'project_type', 'VARCHAR(100)'),
        ('bid_project', 'industry_tags', 'TEXT'),
        ('bid_section', 'section_order', 'INTEGER DEFAULT 0'),
        ('bid_section', 'section_type', 'VARCHAR(50)'),
        ('bid_section', 'source', "VARCHAR(20) DEFAULT 'manual'"),
        ('bid_section', 'original_requirement', 'TEXT'),
        ('bid_section', 'match_score', 'FLOAT DEFAULT 0'),
        ('bid_project', 'tender_requirements', 'TEXT'),
        ('bid_project', 'tender_format', 'TEXT'),
        ('bid_project', 'format_file_path', 'VARCHAR(500)'),
        ('bid_project', 'requirements_file_path', 'VARCHAR(500)'),
        ('bid_section', 'requirement_text', 'TEXT'),
        ('bid_section', 'format_heading_text', 'VARCHAR(500)'),
        ('bid_section', 'format_para_index', 'INTEGER'),
    ]
    for table, col, col_type in _migrate_columns:
        try:
            db.session.execute(db.text(f'ALTER TABLE {table} ADD COLUMN {col} {col_type}'))
        except Exception:
            pass
    db.session.commit()

    # --- 数据修复：修复缺失的 section_type / section_order / bid 状态 ---
    from tender_parser import classify_section
    _fix_count = 0
    for sec in BidSection.query.filter(
        db.or_(BidSection.section_type.is_(None), BidSection.section_type == '')
    ).all():
        sec_type, _cat, _score = classify_section(sec.section_name)
        sec.section_type = sec_type
        sec.match_score = _score
        _fix_count += 1
    if _fix_count:
        db.session.commit()

    # 修复 section_order 全为 0 的项目
    for bid in BidProject.query.all():
        sections = BidSection.query.filter_by(bid_project_id=bid.id)\
            .order_by(BidSection.id).all()
        needs_fix = all(s.section_order == 0 for s in sections) if sections else False
        if needs_fix:
            for i, s in enumerate(sections):
                s.section_order = i + 1
        # 修复卡住的状态
        if sections and bid.status in ('draft', 'parsing'):
            bid.status = 'in_progress'
    db.session.commit()

    # 修复附件分类（授权文件 → 授权委托书）
    CompanyAttachment.query.filter_by(category='授权文件').update(
        {'category': '授权委托书'})
    db.session.commit()


# =============================================================================
# Auth
# =============================================================================

ADMIN_USER = 'admin'
ADMIN_PASS = 'admin123'


def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('logged_in'):
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated


@app.route('/login', methods=['GET', 'POST'])
def login():
    if session.get('logged_in'):
        return redirect(url_for('index'))
    if request.method == 'POST':
        username = request.form.get('username', '')
        password = request.form.get('password', '')
        if username == ADMIN_USER and password == ADMIN_PASS:
            session['logged_in'] = True
            session['username'] = username
            return redirect(url_for('index'))
        flash('用户名或密码错误', 'danger')
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


@app.before_request
def check_login():
    allowed = ('login', 'static')
    if request.endpoint and request.endpoint not in allowed and not session.get('logged_in'):
        return redirect(url_for('login'))


# =============================================================================
# Dashboard (simplified)
# =============================================================================

@app.route('/')
def index():
    bid_count = BidProject.query.count()
    template_count = CompanyAttachment.query.count()
    personnel_count = Personnel.query.count()
    recent_bids = BidProject.query.order_by(BidProject.id.desc()).limit(10).all()
    return render_template('index.html',
                           bid_count=bid_count,
                           template_count=template_count,
                           personnel_count=personnel_count,
                           recent_bids=recent_bids)


# =============================================================================
# Personnel CRUD (simplified form fields)
# =============================================================================

@app.route('/personnel/')
def personnel_list():
    q = request.args.get('q')
    if q:
        s = f'%{q}%'
        personnel = Personnel.query.filter(
            db.or_(Personnel.name.ilike(s), Personnel.title.ilike(s),
                   Personnel.skills.ilike(s), Personnel.position.ilike(s))
        ).all()
    else:
        personnel = Personnel.query.order_by(Personnel.id.desc()).all()
    return render_template('personnel_list.html', personnel=personnel, q=q)


@app.route('/personnel/add', methods=['GET', 'POST'])
def personnel_add():
    if request.method == 'POST':
        p = Personnel(
            name=request.form.get('name'),
            gender=request.form.get('gender'),
            phone=request.form.get('phone'),
            title=request.form.get('title'),
            position=request.form.get('position'),
            skills=request.form.get('skills'),
        )
        db.session.add(p)
        db.session.commit()
        flash('人员添加成功！', 'success')
        return redirect(url_for('personnel_list'))
    return render_template('personnel_form.html', personnel=None)


@app.route('/personnel/<int:id>')
def personnel_detail(id):
    person = Personnel.query.get_or_404(id)
    return render_template('personnel_detail.html', person=person)


@app.route('/personnel/edit/<int:id>', methods=['GET', 'POST'])
def personnel_edit(id):
    person = Personnel.query.get_or_404(id)
    if request.method == 'POST':
        person.name = request.form.get('name')
        person.gender = request.form.get('gender')
        person.phone = request.form.get('phone')
        person.title = request.form.get('title')
        person.position = request.form.get('position')
        person.skills = request.form.get('skills')
        db.session.commit()
        flash('人员信息更新成功！', 'success')
        return redirect(url_for('personnel_detail', id=person.id))
    return render_template('personnel_form.html', personnel=person)


@app.route('/personnel/delete/<int:id>', methods=['POST'])
def personnel_delete(id):
    person = Personnel.query.get_or_404(id)
    for cert in person.certificates:
        if cert.file_path:
            try:
                os.remove(os.path.join(app.config['UPLOAD_FOLDER'], cert.file_path))
            except OSError:
                pass
    db.session.delete(person)
    db.session.commit()
    flash('人员已删除！', 'success')
    return redirect(url_for('personnel_list'))


@app.route('/personnel/<int:id>/cert/add', methods=['POST'])
def personnel_cert_add(id):
    person = Personnel.query.get_or_404(id)
    file = request.files.get('cert_file')
    file_path = save_upload(file, 'personnel')
    cert = Certificate(
        personnel_id=person.id,
        cert_type=request.form.get('cert_type'),
        cert_name=request.form.get('cert_name'),
        cert_number=request.form.get('cert_number'),
        issue_date=parse_date(request.form.get('issue_date')),
        expiry_date=parse_date(request.form.get('expiry_date')),
        file_path=file_path,
    )
    db.session.add(cert)
    db.session.commit()
    flash('证书添加成功！', 'success')
    return redirect(url_for('personnel_edit', id=person.id))


@app.route('/personnel/cert/delete/<int:cert_id>', methods=['POST'])
def certificate_delete(cert_id):
    cert = Certificate.query.get_or_404(cert_id)
    pid = cert.personnel_id
    if cert.file_path:
        try:
            os.remove(os.path.join(app.config['UPLOAD_FOLDER'], cert.file_path))
        except OSError:
            pass
    db.session.delete(cert)
    db.session.commit()
    flash('证书已删除！', 'success')
    return redirect(url_for('personnel_edit', id=pid))


# =============================================================================
# Performance CRUD
# =============================================================================

@app.route('/performance/')
def performance_list():
    q = request.args.get('q')
    if q:
        s = f'%{q}%'
        performances = Performance.query.filter(
            db.or_(Performance.project_name.ilike(s), Performance.client_name.ilike(s),
                   Performance.service_types.ilike(s), Performance.testing_params.ilike(s))
        ).all()
    else:
        performances = Performance.query.order_by(Performance.id.desc()).all()
    return render_template('performance_list.html', performances=performances, q=q)


@app.route('/performance/add', methods=['GET', 'POST'])
def performance_add():
    if request.method == 'POST':
        perf = Performance(
            project_name=request.form.get('project_name'),
            client_name=request.form.get('client_name'),
            contract_amount=parse_float(request.form.get('contract_amount')),
            service_start=parse_date(request.form.get('service_start')),
            service_end=parse_date(request.form.get('service_end')),
            service_types=request.form.get('service_types'),
            testing_params=request.form.get('testing_params'),
            description=request.form.get('description'),
        )
        db.session.add(perf)
        db.session.commit()
        flash('业绩添加成功！', 'success')
        return redirect(url_for('performance_list'))
    return render_template('performance_form.html', performance=None)


@app.route('/performance/<int:id>')
def performance_detail(id):
    perf = Performance.query.get_or_404(id)
    return render_template('performance_detail.html', perf=perf)


@app.route('/performance/edit/<int:id>', methods=['GET', 'POST'])
def performance_edit(id):
    perf = Performance.query.get_or_404(id)
    if request.method == 'POST':
        perf.project_name = request.form.get('project_name')
        perf.client_name = request.form.get('client_name')
        perf.contract_amount = parse_float(request.form.get('contract_amount'))
        perf.service_start = parse_date(request.form.get('service_start'))
        perf.service_end = parse_date(request.form.get('service_end'))
        perf.service_types = request.form.get('service_types')
        perf.testing_params = request.form.get('testing_params')
        perf.description = request.form.get('description')
        db.session.commit()
        flash('业绩信息更新成功！', 'success')
        return redirect(url_for('performance_detail', id=perf.id))
    return render_template('performance_form.html', performance=perf)


@app.route('/performance/delete/<int:id>', methods=['POST'])
def performance_delete(id):
    perf = Performance.query.get_or_404(id)
    for f in perf.files:
        if f.file_path:
            try:
                os.remove(os.path.join(app.config['UPLOAD_FOLDER'], f.file_path))
            except OSError:
                pass
    db.session.delete(perf)
    db.session.commit()
    flash('业绩已删除！', 'success')
    return redirect(url_for('performance_list'))


@app.route('/performance/<int:id>/file/add', methods=['POST'])
def performance_file_add(id):
    perf = Performance.query.get_or_404(id)
    file = request.files.get('file')
    file_path = save_upload(file, 'performance')
    if file_path:
        pf = PerformanceFile(
            performance_id=perf.id,
            file_type=request.form.get('file_type'),
            file_name=file.filename,
            file_path=file_path,
        )
        db.session.add(pf)
        db.session.commit()
        flash('文件上传成功！', 'success')
    else:
        flash('文件上传失败，请检查文件格式！', 'danger')
    return redirect(url_for('performance_edit', id=perf.id))


@app.route('/performance/file/delete/<int:file_id>', methods=['POST'])
def performance_file_delete(file_id):
    pf = PerformanceFile.query.get_or_404(file_id)
    pid = pf.performance_id
    if pf.file_path:
        try:
            os.remove(os.path.join(app.config['UPLOAD_FOLDER'], pf.file_path))
        except OSError:
            pass
    db.session.delete(pf)
    db.session.commit()
    flash('文件已删除！', 'success')
    return redirect(url_for('performance_edit', id=pid))


# =============================================================================
# Attachments (资料模板库) CRUD
# =============================================================================

@app.route('/attachments/')
def attachment_list():
    category = request.args.get('category')
    if category:
        attachments = CompanyAttachment.query.filter_by(category=category).order_by(CompanyAttachment.id.desc()).all()
    else:
        attachments = CompanyAttachment.query.order_by(CompanyAttachment.id.desc()).all()
    categories = [r[0] for r in db.session.query(CompanyAttachment.category).distinct().all() if r[0]]
    return render_template('attachment_list.html', attachments=attachments,
                           categories=categories, current_category=category,
                           now=date.today(),
                           now_plus_90=date.today() + timedelta(days=90))


@app.route('/attachments/add', methods=['GET', 'POST'])
def attachment_add():
    if request.method == 'POST':
        file = request.files.get('file')
        file_path = save_upload(file, 'attachments')
        att = CompanyAttachment(
            name=request.form.get('name'),
            category=request.form.get('category'),
            file_path=file_path,
            issue_date=parse_date(request.form.get('issue_date')),
            expiry_date=parse_date(request.form.get('expiry_date')),
            version=request.form.get('version'),
            tags=request.form.get('tags'),
            notes=request.form.get('notes'),
        )
        db.session.add(att)
        db.session.commit()
        flash('模板添加成功！', 'success')
        return redirect(url_for('attachment_list'))
    return render_template('attachment_form.html', attachment=None)


@app.route('/attachments/edit/<int:id>', methods=['GET', 'POST'])
def attachment_edit(id):
    att = CompanyAttachment.query.get_or_404(id)
    if request.method == 'POST':
        att.name = request.form.get('name')
        att.category = request.form.get('category')
        att.issue_date = parse_date(request.form.get('issue_date'))
        att.expiry_date = parse_date(request.form.get('expiry_date'))
        att.version = request.form.get('version')
        att.tags = request.form.get('tags')
        att.notes = request.form.get('notes')
        file = request.files.get('file')
        if file and file.filename:
            new_path = save_upload(file, 'attachments')
            if new_path:
                if att.file_path:
                    try:
                        os.remove(os.path.join(app.config['UPLOAD_FOLDER'], att.file_path))
                    except OSError:
                        pass
                att.file_path = new_path
        db.session.commit()
        flash('模板更新成功！', 'success')
        return redirect(url_for('attachment_list'))
    return render_template('attachment_form.html', attachment=att)


@app.route('/attachments/delete/<int:id>', methods=['POST'])
def attachment_delete(id):
    att = CompanyAttachment.query.get_or_404(id)
    if att.file_path:
        try:
            os.remove(os.path.join(app.config['UPLOAD_FOLDER'], att.file_path))
        except OSError:
            pass
    db.session.delete(att)
    db.session.commit()
    flash('模板已删除！', 'success')
    return redirect(url_for('attachment_list'))


@app.route('/attachments/download/<int:id>')
def attachment_download(id):
    att = CompanyAttachment.query.get_or_404(id)
    if att.file_path:
        full_path = os.path.join(app.config['UPLOAD_FOLDER'], att.file_path)
        return send_file(full_path, as_attachment=True)
    flash('文件不存在！', 'danger')
    return redirect(url_for('attachment_list'))


# =============================================================================
# Bid Projects CRUD
# =============================================================================

@app.route('/bids/')
def bid_list():
    bids = BidProject.query.order_by(BidProject.id.desc()).all()
    return render_template('bid_list.html', bids=bids)


@app.route('/bids/add', methods=['GET', 'POST'])
def bid_add():
    if request.method == 'POST':
        # 保存格式模板文件
        format_file = request.files.get('format_file')
        format_path = save_upload(format_file, 'attachments') if format_file and format_file.filename else None

        # 保存要求文档文件
        requirements_file = request.files.get('requirements_file')
        requirements_path = save_upload(requirements_file, 'attachments') if requirements_file and requirements_file.filename else None

        # 兼容旧的单文件上传
        tender_file = request.files.get('tender_file')
        tender_path = save_upload(tender_file, 'attachments') if tender_file and tender_file.filename else None

        bid = BidProject(
            project_name=request.form.get('project_name'),
            bidder_name=request.form.get('bidder_name'),
            agent_name=request.form.get('agent_name'),
            max_price=parse_float(request.form.get('max_price')),
            deadline=parse_datetime(request.form.get('deadline')),
            service_period=request.form.get('service_period'),
            project_type=request.form.get('project_type'),
            industry_tags=request.form.get('industry_tags'),
            tender_file_path=tender_path,
            format_file_path=format_path,
            requirements_file_path=requirements_path,
            tender_requirements=request.form.get('tender_requirements', '').strip(),
            tender_format=request.form.get('tender_format', '').strip(),
            status='parsing',
            notes=request.form.get('notes'),
        )
        db.session.add(bid)
        db.session.commit()

        # 新流程：双文件模式（格式模板 + 要求文档）
        if format_path:
            from format_parser import parse_format_template
            fmt_full = os.path.join(app.config['UPLOAD_FOLDER'], format_path)
            fmt_result = parse_format_template(fmt_full)

            # 解析要求文档
            req_result = None
            section_requirements = {}
            if requirements_path:
                req_full = os.path.join(app.config['UPLOAD_FOLDER'], requirements_path)
                req_result = parse_requirements_file(req_full)
                if req_result.get('success'):
                    bid.tender_requirements = req_result['full_text']
                    db.session.commit()

            if fmt_result['success']:
                # 将格式模板解析结果与要求匹配
                if req_result and req_result.get('success'):
                    section_requirements = match_requirements_to_sections(
                        req_result, fmt_result['sections'])

                # 构建章节数据（含格式模板定位信息）
                parsed_sections = []
                for sec in fmt_result['sections']:
                    req_text = section_requirements.get(sec['section_name'], '') or \
                               section_requirements.get(sec['section_type'], '')
                    parsed_sections.append({
                        'order': sec['order'],
                        'section_name': sec['section_name'],
                        'section_type': sec['section_type'],
                        'category': sec['category'],
                        'match_score': sec['match_score'],
                        'original_text': sec['heading_text'],
                        'requirement_text': req_text,
                        'format_heading_text': sec['heading_text'],
                        'format_para_index': sec['para_index'],
                    })

                session['parsed_sections'] = parsed_sections
                session['parsed_context'] = ''
                session['parsed_message'] = fmt_result['message']
                session['parsed_requirements'] = bid.tender_requirements or ''
                session['section_requirements'] = section_requirements
                flash(f'格式模板解析成功：{fmt_result["message"]}', 'success')
                return redirect(url_for('bid_confirm_sections', id=bid.id))
            else:
                flash(f'格式模板解析提示：{fmt_result["message"]}，已加载默认章节', 'warning')

        # 兼容旧流程：单文件模式
        elif tender_path:
            ext = tender_path.rsplit('.', 1)[-1].lower() if '.' in tender_path else ''
            if ext in TENDER_EXTENSIONS:
                full_path = os.path.join(app.config['UPLOAD_FOLDER'], tender_path)
                result = parse_tender_file_dual(full_path)
                if result['success']:
                    if not bid.tender_requirements and result.get('requirements'):
                        bid.tender_requirements = result['requirements']
                    if not bid.tender_format and result.get('format_text'):
                        bid.tender_format = result['format_text']
                    db.session.commit()

                    session['parsed_sections'] = result['sections']
                    session['parsed_context'] = result.get('raw_context', '')
                    session['parsed_message'] = result['message']
                    session['parsed_requirements'] = result.get('requirements', '')
                    session['section_requirements'] = result.get('section_requirements', {})
                    flash(f'招标文件解析成功：{result["message"]}', 'success')
                    return redirect(url_for('bid_confirm_sections', id=bid.id))
                else:
                    flash(f'招标文件解析提示：{result["message"]}，已加载默认章节', 'warning')

        # 解析失败或未上传文件，使用默认章节
        default_secs = get_default_sections()
        session['parsed_sections'] = default_secs
        session['parsed_context'] = ''
        session['parsed_message'] = '使用默认章节结构'
        session['parsed_requirements'] = ''
        session['section_requirements'] = {}
        return redirect(url_for('bid_confirm_sections', id=bid.id))

    return render_template('bid_form.html', bid=None)


@app.route('/bids/<int:id>/confirm-sections', methods=['GET', 'POST'])
def bid_confirm_sections(id):
    """解析结果确认页面：用户可编辑、增删、排序章节列表后确认"""
    bid = BidProject.query.get_or_404(id)

    if request.method == 'POST':
        # 删除旧章节
        BidSection.query.filter_by(bid_project_id=bid.id).delete()

        # 从表单读取确认后的章节
        names = request.form.getlist('section_name')
        types = request.form.getlist('section_type')
        originals = request.form.getlist('original_text')
        req_texts = request.form.getlist('requirement_text')
        heading_texts = request.form.getlist('format_heading_text')
        para_indices = request.form.getlist('format_para_index')

        for i, name in enumerate(names):
            name = name.strip()
            if not name:
                continue
            para_idx = None
            if i < len(para_indices) and para_indices[i]:
                try:
                    para_idx = int(para_indices[i])
                except (ValueError, TypeError):
                    pass
            sec = BidSection(
                bid_project_id=bid.id,
                section_key=f'section_{i+1}',
                section_name=name,
                section_order=i + 1,
                section_type=types[i] if i < len(types) else '其他',
                source='parsed' if (originals[i] if i < len(originals) else '') else 'manual',
                original_requirement=originals[i] if i < len(originals) else '',
                requirement_text=req_texts[i] if i < len(req_texts) else '',
                format_heading_text=heading_texts[i] if i < len(heading_texts) else '',
                format_para_index=para_idx,
                status='pending',
            )
            db.session.add(sec)

        bid.status = 'in_progress'
        db.session.commit()

        # 自动匹配
        stats = auto_match_all_sections(bid)
        flash(f'已确认 {len(names)} 个章节，自动匹配了 {stats["matched_sections"]} 个资料模板', 'success')

        # 清理 session
        session.pop('parsed_sections', None)
        session.pop('parsed_context', None)
        session.pop('parsed_message', None)
        session.pop('parsed_requirements', None)
        session.pop('section_requirements', None)

        return redirect(url_for('bid_detail', id=bid.id))

    # GET: 显示解析结果供确认
    sections = session.get('parsed_sections', get_default_sections())
    raw_context = session.get('parsed_context', '')
    message = session.get('parsed_message', '')
    section_requirements = session.get('section_requirements', {})

    # 将匹配的要求注入到各章节数据中
    for sec in sections:
        sec_name = sec.get('section_name', '')
        sec_type = sec.get('section_type', '')
        # 按章节名或类型查找匹配的要求
        req = section_requirements.get(sec_name, '') or section_requirements.get(sec_type, '')
        sec['requirement_text'] = req

    # 收集所有可用的 section_type 供下拉选择
    from tender_parser import SECTION_KEYWORD_MAP
    available_types = sorted(SECTION_KEYWORD_MAP.keys())

    return render_template('bid_confirm_sections.html',
                           bid=bid,
                           sections=sections,
                           raw_context=raw_context,
                           message=message,
                           available_types=available_types)


@app.route('/bids/<int:id>/reparse', methods=['POST'])
def bid_reparse(id):
    """重新解析招标文件（优先格式模板，兼容旧单文件）"""
    bid = BidProject.query.get_or_404(id)

    # 新流程：格式模板重新解析
    if bid.format_file_path:
        from format_parser import parse_format_template
        fmt_full = os.path.join(app.config['UPLOAD_FOLDER'], bid.format_file_path)
        fmt_result = parse_format_template(fmt_full)

        req_result = None
        section_requirements = {}
        if bid.requirements_file_path:
            req_full = os.path.join(app.config['UPLOAD_FOLDER'], bid.requirements_file_path)
            req_result = parse_requirements_file(req_full)
            if req_result.get('success'):
                bid.tender_requirements = req_result['full_text']
                db.session.commit()

        if fmt_result['success']:
            if req_result and req_result.get('success'):
                section_requirements = match_requirements_to_sections(
                    req_result, fmt_result['sections'])

            parsed_sections = []
            for sec in fmt_result['sections']:
                req_text = section_requirements.get(sec['section_name'], '') or \
                           section_requirements.get(sec['section_type'], '')
                parsed_sections.append({
                    'order': sec['order'],
                    'section_name': sec['section_name'],
                    'section_type': sec['section_type'],
                    'category': sec['category'],
                    'match_score': sec['match_score'],
                    'original_text': sec['heading_text'],
                    'requirement_text': req_text,
                    'format_heading_text': sec['heading_text'],
                    'format_para_index': sec['para_index'],
                })

            session['parsed_sections'] = parsed_sections
            session['parsed_context'] = ''
            session['parsed_message'] = fmt_result['message']
            session['parsed_requirements'] = bid.tender_requirements or ''
            session['section_requirements'] = section_requirements
            flash(f'重新解析成功：{fmt_result["message"]}', 'success')
            return redirect(url_for('bid_confirm_sections', id=bid.id))
        else:
            flash(f'格式模板解析失败：{fmt_result["message"]}', 'warning')

    # 旧流程：单文件重新解析
    if bid.tender_file_path:
        full_path = os.path.join(app.config['UPLOAD_FOLDER'], bid.tender_file_path)
        result = parse_tender_file_dual(full_path)
        if result['success']:
            if result.get('requirements'):
                bid.tender_requirements = result['requirements']
            if result.get('format_text'):
                bid.tender_format = result['format_text']
            db.session.commit()

            session['parsed_sections'] = result['sections']
            session['parsed_context'] = result.get('raw_context', '')
            session['parsed_message'] = result['message']
            session['parsed_requirements'] = result.get('requirements', '')
            session['section_requirements'] = result.get('section_requirements', {})
            flash(f'重新解析成功：{result["message"]}', 'success')
            return redirect(url_for('bid_confirm_sections', id=bid.id))
        else:
            flash(f'解析失败：{result["message"]}', 'warning')

    # 都失败了
    if not bid.format_file_path and not bid.tender_file_path:
        flash('未上传格式模板或招标文件，无法解析', 'danger')
    session['parsed_sections'] = get_default_sections()
    session['parsed_context'] = ''
    session['parsed_message'] = '使用默认章节结构'
    session['section_requirements'] = {}
    return redirect(url_for('bid_confirm_sections', id=bid.id))


@app.route('/bids/<int:id>/auto-match', methods=['POST'])
def bid_auto_match(id):
    """手动触发自动匹配"""
    bid = BidProject.query.get_or_404(id)
    stats = auto_match_all_sections(bid)
    flash(f'自动匹配完成：{stats["matched_sections"]}/{stats["total_sections"]} 个章节已匹配资料模板', 'success')
    return redirect(url_for('bid_detail', id=bid.id))


@app.route('/bids/<int:id>')
def bid_detail(id):
    bid = BidProject.query.get_or_404(id)
    # 按 section_order 排序
    sections = BidSection.query.filter_by(bid_project_id=bid.id)\
        .order_by(BidSection.section_order).all()
    total_sections = len(sections)
    done_sections = sum(1 for s in sections if s.status == 'done')
    return render_template('bid_detail.html', bid=bid,
                           sections=sections,
                           total_sections=total_sections,
                           done_sections=done_sections)


@app.route('/bids/edit/<int:id>', methods=['GET', 'POST'])
def bid_edit(id):
    bid = BidProject.query.get_or_404(id)
    if request.method == 'POST':
        bid.project_name = request.form.get('project_name')
        bid.bidder_name = request.form.get('bidder_name')
        bid.agent_name = request.form.get('agent_name')
        bid.max_price = parse_float(request.form.get('max_price'))
        bid.deadline = parse_datetime(request.form.get('deadline'))
        bid.service_period = request.form.get('service_period')
        bid.project_type = request.form.get('project_type')
        bid.industry_tags = request.form.get('industry_tags')
        bid.tender_requirements = request.form.get('tender_requirements', '').strip()
        bid.tender_format = request.form.get('tender_format', '').strip()
        bid.notes = request.form.get('notes')

        # 更新格式模板文件
        format_file = request.files.get('format_file')
        if format_file and format_file.filename:
            new_path = save_upload(format_file, 'attachments')
            if new_path:
                if bid.format_file_path:
                    try:
                        os.remove(os.path.join(app.config['UPLOAD_FOLDER'], bid.format_file_path))
                    except OSError:
                        pass
                bid.format_file_path = new_path

        # 更新要求文档文件
        requirements_file = request.files.get('requirements_file')
        if requirements_file and requirements_file.filename:
            new_path = save_upload(requirements_file, 'attachments')
            if new_path:
                if bid.requirements_file_path:
                    try:
                        os.remove(os.path.join(app.config['UPLOAD_FOLDER'], bid.requirements_file_path))
                    except OSError:
                        pass
                bid.requirements_file_path = new_path

        # 兼容旧的单文件上传
        tender_file = request.files.get('tender_file')
        if tender_file and tender_file.filename:
            new_path = save_upload(tender_file, 'attachments')
            if new_path:
                if bid.tender_file_path:
                    try:
                        os.remove(os.path.join(app.config['UPLOAD_FOLDER'], bid.tender_file_path))
                    except OSError:
                        pass
                bid.tender_file_path = new_path
        db.session.commit()
        flash('投标项目更新成功！', 'success')
        return redirect(url_for('bid_detail', id=bid.id))
    return render_template('bid_form.html', bid=bid)


@app.route('/bids/delete/<int:id>', methods=['POST'])
def bid_delete(id):
    bid = BidProject.query.get_or_404(id)
    db.session.delete(bid)
    db.session.commit()
    flash('投标项目已删除！', 'success')
    return redirect(url_for('bid_list'))


@app.route('/bids/<int:bid_id>/section/add', methods=['POST'])
def bid_section_add(bid_id):
    """添加新章节"""
    bid = BidProject.query.get_or_404(bid_id)
    max_order = db.session.query(db.func.max(BidSection.section_order))\
        .filter_by(bid_project_id=bid.id).scalar() or 0
    sec = BidSection(
        bid_project_id=bid.id,
        section_key=f'section_{max_order + 1}',
        section_name=request.form.get('section_name', '新章节'),
        section_order=max_order + 1,
        section_type=request.form.get('section_type', '其他'),
        source='manual',
        status='pending',
    )
    db.session.add(sec)
    db.session.commit()
    flash(f'已添加章节：{sec.section_name}', 'success')
    return redirect(url_for('bid_detail', id=bid.id))


@app.route('/bids/<int:bid_id>/section/<int:section_id>/delete', methods=['POST'])
def bid_section_delete(bid_id, section_id):
    """删除章节"""
    sec = BidSection.query.get_or_404(section_id)
    name = sec.section_name
    db.session.delete(sec)
    db.session.commit()
    flash(f'已删除章节：{name}', 'success')
    return redirect(url_for('bid_detail', id=bid_id))


@app.route('/bids/<int:bid_id>/sections/reorder', methods=['POST'])
def bid_sections_reorder(bid_id):
    """拖拽排序章节"""
    order_data = request.get_json()
    if order_data and 'order' in order_data:
        for i, section_id in enumerate(order_data['order']):
            sec = BidSection.query.get(int(section_id))
            if sec and sec.bid_project_id == bid_id:
                sec.section_order = i + 1
        db.session.commit()
    return jsonify({'success': True})


# =============================================================================
# Bid Section Edit
# =============================================================================

@app.route('/bids/<int:bid_id>/section/<int:section_id>', methods=['GET', 'POST'])
def bid_section_edit(bid_id, section_id):
    bid = BidProject.query.get_or_404(bid_id)
    section = BidSection.query.get_or_404(section_id)

    # 根据 section_type 确定推荐类别
    from tender_parser import SECTION_KEYWORD_MAP
    suggested_category = '其他'
    for stype, info in SECTION_KEYWORD_MAP.items():
        if stype == section.section_type:
            suggested_category = info['category']
            break

    # Get templates filtered by suggested category, plus all templates
    category_templates = CompanyAttachment.query.filter_by(category=suggested_category).order_by(CompanyAttachment.id.desc()).all()
    all_templates = CompanyAttachment.query.order_by(CompanyAttachment.category, CompanyAttachment.id.desc()).all()

    # 获取匹配信息
    personnel_matches = None
    performance_matches = None
    if section.section_type == '人员资料':
        personnel_matches = match_personnel(bid.industry_tags)
    elif section.section_type == '业绩证明':
        performance_matches = match_performance(bid.industry_tags)

    if request.method == 'POST':
        section.section_name = request.form.get('section_name', section.section_name)

        # 处理上传新文件：自动创建附件记录并关联
        upload_file = request.files.get('upload_file')
        if upload_file and upload_file.filename:
            file_path = save_upload(upload_file, 'attachments')
            if file_path:
                upload_name = request.form.get('upload_name', '').strip()
                if not upload_name:
                    upload_name = upload_file.filename
                new_att = CompanyAttachment(
                    name=upload_name,
                    category=suggested_category,
                    file_path=file_path,
                )
                db.session.add(new_att)
                db.session.flush()
                section.attachment_id = new_att.id
                flash(f'已上传并保存到资料模板库：{upload_name}', 'success')
            else:
                flash('文件上传失败，请检查文件格式', 'danger')
        else:
            att_id = request.form.get('attachment_id')
            section.attachment_id = int(att_id) if att_id else None

        section.custom_content = request.form.get('custom_content')
        section.status = request.form.get('status', 'pending')
        db.session.commit()
        flash(f'"{section.section_name}" 已保存', 'success')
        return redirect(url_for('bid_detail', id=bid.id))

    # 找到前后章节用于导航
    all_sections = BidSection.query.filter_by(bid_project_id=bid.id)\
        .order_by(BidSection.section_order).all()
    prev_section = None
    next_section = None
    for i, s in enumerate(all_sections):
        if s.id == section.id:
            if i > 0:
                prev_section = all_sections[i - 1]
            if i < len(all_sections) - 1:
                next_section = all_sections[i + 1]
            break

    return render_template('bid_section_edit.html',
                           bid=bid,
                           section=section,
                           suggested_category=suggested_category,
                           category_templates=category_templates,
                           all_templates=all_templates,
                           personnel_matches=personnel_matches,
                           performance_matches=performance_matches,
                           prev_section=prev_section,
                           next_section=next_section)


# =============================================================================
# Preview / Generate (使用 section_generators 生成专业格式文档)
# =============================================================================

from section_generators import generate_full_bid, generate_section_preview
from template_filler import generate_filled_document
import tempfile


@app.route('/bids/<int:bid_id>/section/<int:section_id>/preview')
@login_required
def bid_section_preview(bid_id, section_id):
    """生成单个章节预览"""
    bid = BidProject.query.get_or_404(bid_id)
    section = BidSection.query.get_or_404(section_id)

    os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)
    doc = generate_section_preview(section, bid, app.config)
    tmp = tempfile.NamedTemporaryFile(suffix='.docx', delete=False, dir=app.config['OUTPUT_FOLDER'])
    doc.save(tmp.name)
    return send_file(tmp.name, as_attachment=True,
                     download_name=f'{section.section_name}_预览.docx')


@app.route('/bids/<int:bid_id>/preview')
@login_required
def bid_preview(bid_id):
    """生成整个投标文件预览（优先使用格式模板填充，否则回退到从零生成）"""
    bid = BidProject.query.get_or_404(bid_id)
    os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

    safe_name = bid.project_name.replace('/', '_')[:50]

    # 新流程：有格式模板文件时，使用模板填充引擎
    if bid.format_file_path:
        fmt_full = os.path.join(app.config['UPLOAD_FOLDER'], bid.format_file_path)
        if os.path.exists(fmt_full):
            sections_db = BidSection.query.filter_by(bid_project_id=bid.id)\
                .order_by(BidSection.section_order).all()
            try:
                doc, output_path = generate_filled_document(
                    fmt_full, bid, sections_db, app.config)
                bid.output_file_path = output_path
                db.session.commit()
                return send_file(output_path, as_attachment=True,
                                 download_name=f'{safe_name}_投标文件.docx')
            except Exception as e:
                flash(f'模板填充失败，回退到标准生成：{e}', 'warning')

    # 旧流程回退：从零生成
    doc = generate_full_bid(bid, app.config)
    tmp = tempfile.NamedTemporaryFile(suffix='.docx', delete=False, dir=app.config['OUTPUT_FOLDER'])
    doc.save(tmp.name)

    bid.output_file_path = tmp.name
    db.session.commit()

    return send_file(tmp.name, as_attachment=True,
                     download_name=f'{safe_name}_投标文件预览.docx')


# =============================================================================
# Static file serving for uploads
# =============================================================================

@app.route('/uploads/<path:filename>')
def uploaded_file(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename))


# =============================================================================
# Run
# =============================================================================

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001, debug=True)
