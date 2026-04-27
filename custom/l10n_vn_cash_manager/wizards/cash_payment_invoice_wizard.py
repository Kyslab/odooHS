from odoo import models, fields, api, _
from odoo.exceptions import UserError


class CashPaymentInvoiceLine(models.TransientModel):
    """Một dòng công nợ phải trả chưa thanh toán trong wizard"""
    _name = 'cash.payment.invoice.line'
    _description = 'Dòng hóa đơn chưa trả'
    _order = 'invoice_date asc, invoice_name asc'

    wizard_id = fields.Many2one(
        'cash.payment.invoice.wizard',
        required=True,
        ondelete='cascade',
    )
    # KHÔNG readonly=True ở model — để browser gửi giá trị khi save TransientModel
    invoice_id = fields.Many2one(
        'account.move',
        string='Chứng từ',
    )
    invoice_name = fields.Char(
        string='Số chứng từ',
        related='invoice_id.name',
        readonly=True,
        store=False,
    )
    invoice_date = fields.Date(
        string='Ngày',
        related='invoice_id.date',
        readonly=True,
        store=False,
    )
    dien_giai = fields.Char(
        string='Diễn giải',
        related='invoice_id.x_dien_giai',
        readonly=True,
        store=False,
    )
    amount_total = fields.Monetary(
        string='Tổng tiền HĐ',
        currency_field='currency_id',
    )
    amount_residual = fields.Monetary(
        string='Còn phải trả',
        currency_field='currency_id',
    )
    currency_id = fields.Many2one(
        'res.currency',
        related='wizard_id.currency_id',
        readonly=True,
    )
    amount_to_pay = fields.Monetary(
        string='Tiền mặt trả',
        currency_field='currency_id',
    )
    selected = fields.Boolean(
        string='Chọn',
        default=True,
    )

    # ── Chiết khấu (được hưởng từ NCC) ──────────────────────────────────
    chiet_khau_amount = fields.Monetary(
        string='Chiết khấu',
        currency_field='currency_id',
        default=0.0,
    )
    chiet_khau_account_id = fields.Many2one(
        'account.account',
        string='TK chiết khấu',
        domain=[('deprecated', '=', False)],
    )
    chiet_khau_analytic_id = fields.Many2one(
        'account.analytic.account',
        string='Khoản mục',
    )

    @api.onchange('chiet_khau_amount')
    def _onchange_chiet_khau_amount(self):
        """Nhập chiết khấu → tự giảm tiền mặt trả = số dư còn lại − chiết khấu"""
        if self.chiet_khau_amount < 0:
            self.chiet_khau_amount = 0.0
        if self.amount_residual:
            self.amount_to_pay = max(0.0, self.amount_residual - self.chiet_khau_amount)

    # ── TK công nợ (dùng nội bộ cho đối trừ) ────────────────────────────
    tk331_account_id = fields.Many2one(
        'account.account',
        string='TK phải trả',
    )
    account_code = fields.Char(
        string='TK',
        related='tk331_account_id.code',
        readonly=True,
        store=False,
    )
    invoice_tk331_line_id = fields.Many2one(
        'account.move.line',
        string='Dòng TK phải trả HĐ',
    )


class CashPaymentInvoiceWizard(models.TransientModel):
    """Wizard trả tiền theo hóa đơn — chọn NCC → load CT chưa trả → trả tiền + đối trừ"""
    _name = 'cash.payment.invoice.wizard'
    _description = 'Trả Tiền Theo Hóa Đơn'

    # ── Thông tin chứng từ ───────────────────────────────────────────────
    date = fields.Date(
        string='Ngày hạch toán',
        required=True,
        default=fields.Date.context_today,
    )
    partner_id = fields.Many2one(
        'res.partner',
        string='Nhà cung cấp',
        required=True,
    )
    journal_id = fields.Many2one(
        'account.journal',
        string='Quỹ tiền mặt',
        required=True,
        domain=[('type', '=', 'cash')],
        default=lambda self: self._default_journal(),
    )
    currency_id = fields.Many2one(
        'res.currency',
        default=lambda self: self.env.company.currency_id,
    )
    ly_do = fields.Char(
        string='Lý do chi tiền',
        default='Trả tiền theo hóa đơn',
    )
    dien_giai = fields.Char(
        string='Diễn giải',
    )

    # ── Danh sách công nợ ────────────────────────────────────────────────
    invoice_line_ids = fields.One2many(
        'cash.payment.invoice.line',
        'wizard_id',
        string='Hóa đơn chưa trả',
    )

    # ── Bộ lọc danh sách ────────────────────────────────────────────────
    f_date_from  = fields.Date(string='Từ ngày')
    f_date_to    = fields.Date(string='Đến ngày')
    f_so_ct      = fields.Char(string='Số CT')
    f_tk         = fields.Char(string='TK')
    f_dien_giai  = fields.Char(string='Lọc diễn giải')

    # ── Tổng cộng (computed) ─────────────────────────────────────────────
    total_tien_mat = fields.Monetary(
        string='Tổng tiền mặt',
        currency_field='currency_id',
        compute='_compute_totals',
        store=False,
    )
    total_chiet_khau = fields.Monetary(
        string='Tổng chiết khấu',
        currency_field='currency_id',
        compute='_compute_totals',
        store=False,
    )
    total_thanh_toan = fields.Monetary(
        string='Tổng thanh toán',
        currency_field='currency_id',
        compute='_compute_totals',
        store=False,
    )
    count_selected = fields.Integer(
        string='Số CT chọn',
        compute='_compute_totals',
        store=False,
    )

    # ── Helpers ─────────────────────────────────────────────────────────
    def _default_journal(self):
        return self.env['account.journal'].search(
            [('type', '=', 'cash'), ('company_id', '=', self.env.company.id)],
            limit=1,
        )

    @api.depends(
        'invoice_line_ids.selected',
        'invoice_line_ids.amount_to_pay',
        'invoice_line_ids.chiet_khau_amount',
    )
    def _compute_totals(self):
        for rec in self:
            selected = rec.invoice_line_ids.filtered('selected')
            rec.total_tien_mat = sum(selected.mapped('amount_to_pay'))
            rec.total_chiet_khau = sum(selected.mapped('chiet_khau_amount'))
            rec.total_thanh_toan = rec.total_tien_mat + rec.total_chiet_khau
            rec.count_selected = len(selected)

    # ── Chọn / bỏ chọn tất cả ───────────────────────────────────────────
    chon_tat_ca = fields.Boolean(
        string='Chọn tất cả',
        default=True,
    )

    @api.onchange('chon_tat_ca')
    def _onchange_chon_tat_ca(self):
        for line in self.invoice_line_ids:
            line.selected = self.chon_tat_ca

    # ── Load công nợ phải trả ────────────────────────────────────────────
    @api.onchange('partner_id')
    def _onchange_partner_id(self):
        self.invoice_line_ids = [(5, 0, 0)]
        if self.partner_id:
            self._load_invoices()

    def _reload_action(self):
        return {
            'type': 'ir.actions.act_window',
            'res_model': self._name,
            'res_id': self.id,
            'view_mode': 'form',
            'target': 'new',
        }

    def action_load_invoices(self):
        """Tải lại toàn bộ — xóa bộ lọc"""
        self.ensure_one()
        if not self.partner_id:
            raise UserError(_('Vui lòng chọn nhà cung cấp trước!'))
        self.write({'f_date_from': False, 'f_date_to': False, 'f_so_ct': False, 'f_tk': False, 'f_dien_giai': False})
        self.invoice_line_ids = [(5, 0, 0)]
        self._load_invoices()
        return self._reload_action()

    def action_apply_filter(self):
        """Áp dụng bộ lọc — reload danh sách theo tiêu chí đã nhập"""
        self.ensure_one()
        if not self.partner_id:
            raise UserError(_('Vui lòng chọn nhà cung cấp trước!'))
        self.invoice_line_ids = [(5, 0, 0)]
        self._load_invoices(
            date_from=self.f_date_from,
            date_to=self.f_date_to,
            so_ct=self.f_so_ct,
            tk=self.f_tk,
            dien_giai=self.f_dien_giai,
        )
        return self._reload_action()

    def action_clear_filter(self):
        """Xóa bộ lọc — load lại toàn bộ"""
        self.ensure_one()
        self.write({'f_date_from': False, 'f_date_to': False, 'f_so_ct': False, 'f_tk': False, 'f_dien_giai': False})
        self.invoice_line_ids = [(5, 0, 0)]
        self._load_invoices()
        return self._reload_action()

    def _load_invoices(self, date_from=False, date_to=False, so_ct=False, tk=False, dien_giai=False):
        """Tìm tất cả dòng công nợ phải trả còn dư Có (TK 331, 3388, ...) của NCC"""
        tk_domain = [
            ('partner_id', '=', self.partner_id.id),
            ('account_id.reconcile', '=', True),
            ('credit', '>', 0),
            ('amount_residual', '<', 0),
            ('parent_state', '=', 'posted'),
            ('reconciled', '=', False),
        ]
        if date_from:
            tk_domain.append(('move_id.date', '>=', str(date_from)))
        if date_to:
            tk_domain.append(('move_id.date', '<=', str(date_to)))
        if so_ct:
            tk_domain.append(('move_id.name', 'ilike', so_ct))
        if tk:
            tk_domain.append(('account_id.code', 'ilike', tk))
        if dien_giai:
            tk_domain.append(('move_id.x_dien_giai', 'ilike', dien_giai))
        tk_lines = self.env['account.move.line'].search(tk_domain)
        tk_lines_sorted = tk_lines.sorted(
            key=lambda l: (l.move_id.date or '', l.move_id.name or '')
        )

        # Tìm TK chiết khấu mặc định (515 — Doanh thu tài chính / CK được hưởng)
        ck_account = self.env['account.account'].search(
            [('code', '=', '515'), ('deprecated', '=', False)], limit=1
        )
        ck_analytic = self.env['account.analytic.account'].search(
            [('code', '=', '7001'), ('plan_id.name', 'ilike', 'thống kê')], limit=1
        ) or self.env['account.analytic.account'].search(
            [('code', '=', '7001'), ('name', 'ilike', 'thu nhập khác')], limit=1
        )

        lines_to_create = []
        seen_moves = set()
        for ml in tk_lines_sorted:
            move = ml.move_id
            if move.id in seen_moves:
                continue
            seen_moves.add(move.id)

            # amount_residual âm → lấy giá trị tuyệt đối để hiển thị
            residual = abs(ml.amount_residual)
            lines_to_create.append((0, 0, {
                'invoice_id': move.id,
                'amount_total': move.amount_total,
                'amount_residual': residual,
                'amount_to_pay': residual,
                'selected': True,
                'tk331_account_id': ml.account_id.id,
                'invoice_tk331_line_id': ml.id,
                'chiet_khau_amount': 0.0,
                'chiet_khau_account_id': ck_account.id if ck_account else False,
                'chiet_khau_analytic_id': ck_analytic.id if ck_analytic else False,
            }))

        self.invoice_line_ids = lines_to_create

    # ── Xác nhận trả tiền ────────────────────────────────────────────────
    def action_confirm(self):
        self.ensure_one()

        selected_lines = self.invoice_line_ids.filtered(
            lambda l: l.selected and (l.amount_to_pay > 0 or l.chiet_khau_amount > 0)
        )
        if not selected_lines:
            raise UserError(_('Vui lòng chọn ít nhất một chứng từ!'))

        # Kiểm tra số tiền không vượt số dư thực tế
        for ln in selected_lines:
            inv_name = ln.invoice_id.name if ln.invoice_id else '?'
            settle = ln.amount_to_pay + ln.chiet_khau_amount
            actual_residual = abs(ln.invoice_tk331_line_id.amount_residual) \
                if ln.invoice_tk331_line_id else 0.0
            if settle <= 0:
                raise UserError(_(
                    'Tổng tiền mặt + chiết khấu của %s phải lớn hơn 0!', inv_name,
                ))
            if settle > actual_residual + 0.01:
                raise UserError(_(
                    'Tổng thanh toán (%(settle)s) vượt quá số còn phải trả (%(res)s) của %(inv)s!',
                    settle=settle, res=actual_residual, inv=inv_name,
                ))
            if ln.chiet_khau_amount > 0 and not ln.chiet_khau_account_id:
                raise UserError(_(
                    'Vui lòng chọn TK chiết khấu cho %s!', inv_name,
                ))

        # TK tiền mặt của journal
        cash_account = self.journal_id.default_account_id
        if not cash_account:
            raise UserError(_(
                'Quỹ tiền mặt "%s" chưa có tài khoản mặc định. '
                'Vào Cấu hình → Nhật ký để thiết lập.' % self.journal_id.name
            ))

        dien_giai = self.dien_giai or self.ly_do or 'Trả tiền theo hóa đơn'

        # ── Chuẩn bị reconcile pairs ────────────────────────────────────
        reconcile_pairs = []
        total_cash = 0.0
        for ln in selected_lines:
            inv_name = ln.invoice_id.name if ln.invoice_id else ''
            unique_name = '%s [HĐ:%d]' % (inv_name, ln.invoice_id.id)
            settle = ln.amount_to_pay + ln.chiet_khau_amount
            total_cash += ln.amount_to_pay
            reconcile_pairs.append({
                'invoice_cr_line': ln.invoice_tk331_line_id,   # dòng Có TK 331 gốc
                'account_id': ln.tk331_account_id.id,
                'settle_amount': settle,
                'cash_amount': ln.amount_to_pay,
                'chiet_khau_amount': ln.chiet_khau_amount,
                'chiet_khau_account_id': ln.chiet_khau_account_id.id if ln.chiet_khau_account_id else False,
                'chiet_khau_analytic_id': ln.chiet_khau_analytic_id.id if ln.chiet_khau_analytic_id else False,
                'invoice_name': inv_name,
                'line_name': unique_name,
            })

        # ── Tạo bút toán ────────────────────────────────────────────────
        # Trả tiền NCC:
        #   Nợ TK 331/3388  : settle_amount (đối trừ công nợ)
        #   Có TK 515       : chiet_khau_amount (chiết khấu được hưởng)
        #   Có TK tiền mặt  : total_cash
        move_lines = [
            # Có TK tiền mặt — phần thực trả bằng tiền mặt
            (0, 0, {
                'account_id': cash_account.id,
                'partner_id': self.partner_id.id,
                'debit': 0.0,
                'credit': total_cash,
                'name': dien_giai,
            }),
        ]

        for pair in reconcile_pairs:
            # Có TK 515 — phần chiết khấu (nếu có)
            if pair['chiet_khau_amount'] > 0 and pair['chiet_khau_account_id']:
                ck_line = {
                    'account_id': pair['chiet_khau_account_id'],
                    'partner_id': self.partner_id.id,
                    'debit': 0.0,
                    'credit': pair['chiet_khau_amount'],
                    'name': 'CK - %s' % pair['invoice_name'],
                }
                if pair['chiet_khau_analytic_id']:
                    ck_line['analytic_distribution'] = {
                        str(pair['chiet_khau_analytic_id']): 100
                    }
                move_lines.append((0, 0, ck_line))

            # Nợ TK 331/3388 — toàn bộ số thanh toán (tiền mặt + chiết khấu)
            move_lines.append((0, 0, {
                'account_id': pair['account_id'],
                'partner_id': self.partner_id.id,
                'debit': pair['settle_amount'],
                'credit': 0.0,
                'name': pair['line_name'],
            }))

        move_vals = {
            'move_type': 'entry',
            'date': self.date,
            'journal_id': self.journal_id.id,
            'partner_id': self.partner_id.id,
            'ref': dien_giai,
            'x_ly_do': self.ly_do,
            'line_ids': move_lines,
        }

        move = self.env['account.move'].create(move_vals)
        move.action_post()

        # ── Đối trừ công nợ ─────────────────────────────────────────────
        reconcile_errors = []
        for pair in reconcile_pairs:
            invoice_cr_line = pair['invoice_cr_line']
            if not invoice_cr_line:
                reconcile_errors.append(pair['invoice_name'] or '?')
                continue
            # Tìm dòng Nợ TK 331 trong bút toán vừa tạo
            new_dr_line = move.line_ids.filtered(
                lambda l, name=pair['line_name']: l.debit > 0 and l.name == name
            )
            if not new_dr_line:
                reconcile_errors.append(pair['invoice_name'] or '?')
                continue
            try:
                (invoice_cr_line | new_dr_line[0]).reconcile()
            except Exception as e:
                reconcile_errors.append('%s: %s' % (pair['invoice_name'], str(e)))

        # ── Mở phiếu chi ────────────────────────────────────────────────
        result = {
            'type': 'ir.actions.act_window',
            'name': _('Phiếu Chi'),
            'res_model': 'account.move',
            'res_id': move.id,
            'view_mode': 'form',
            'target': 'current',
        }

        if reconcile_errors:
            return {
                'type': 'ir.actions.client',
                'tag': 'display_notification',
                'params': {
                    'title': _('Phiếu chi đã tạo — Lưu ý đối trừ'),
                    'message': _(
                        'Phiếu chi đã tạo thành công nhưng không thể đối trừ tự động cho: %s. '
                        'Vui lòng đối trừ thủ công.'
                    ) % ', '.join(reconcile_errors),
                    'type': 'warning',
                    'next': result,
                },
            }

        return result
