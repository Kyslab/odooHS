from odoo import api, fields, models, _
from odoo.exceptions import UserError


class DebtOffsetWizardLine(models.TransientModel):
    """Dòng công nợ trong wizard bù trừ 131 vs 331"""
    _name = 'debt.offset.wizard.line'
    _description = 'Dòng bù trừ công nợ'

    wizard_id = fields.Many2one('debt.offset.wizard', ondelete='cascade')
    side = fields.Selection([
        ('receivable', 'Phải thu (131)'),
        ('payable',    'Phải trả (331)'),
    ])
    selected = fields.Boolean(string='Chọn', default=False)
    move_line_id = fields.Many2one('account.move.line', readonly=True)

    date = fields.Date(related='move_line_id.date', string='Ngày', readonly=True)
    move_name = fields.Char(related='move_line_id.move_id.name', string='Số CT', readonly=True)
    ref = fields.Char(related='move_line_id.ref', string='Diễn giải', readonly=True)
    amount_residual = fields.Monetary(
        related='move_line_id.amount_residual', string='Số dư còn lại', readonly=True)
    currency_id = fields.Many2one(
        related='move_line_id.company_currency_id', readonly=True)


class DebtOffsetWizard(models.TransientModel):
    """Wizard bù trừ công nợ phải thu (TK131) vs phải trả (TK331) cùng đối tác"""
    _name = 'debt.offset.wizard'
    _description = 'Bù trừ công nợ 131 vs 331'

    partner_id = fields.Many2one(
        'res.partner', string='Đối tác', required=True,
        help='Đối tác vừa là khách hàng (TK131) vừa là nhà cung cấp (TK331)')
    journal_id = fields.Many2one(
        'account.journal', string='Nhật ký bù trừ', required=True,
        domain=[('type', '=', 'general')],
        help='Nhật ký dùng để tạo bút toán bù trừ (thường là nhật ký Bút toán điều chỉnh)')
    date = fields.Date(
        string='Ngày bù trừ', required=True, default=fields.Date.today)
    ref = fields.Char(
        string='Diễn giải', default='Bù trừ công nợ phải thu / phải trả')

    # Danh sách phải thu (TK 131 dư Nợ)
    receivable_line_ids = fields.One2many(
        'debt.offset.wizard.line', 'wizard_id',
        string='Công nợ phải thu (TK 131)',
        domain=[('side', '=', 'receivable')])

    # Danh sách phải trả (TK 331 dư Có)
    payable_line_ids = fields.One2many(
        'debt.offset.wizard.line', 'wizard_id',
        string='Công nợ phải trả (TK 331)',
        domain=[('side', '=', 'payable')])

    currency_id = fields.Many2one(
        'res.currency',
        default=lambda self: self.env.company.currency_id)

    # Tổng hợp
    total_receivable_selected = fields.Monetary(
        string='Tổng phải thu chọn', compute='_compute_totals')
    total_payable_selected = fields.Monetary(
        string='Tổng phải trả chọn', compute='_compute_totals')
    amount_to_offset = fields.Monetary(
        string='Số tiền bù trừ', compute='_compute_totals',
        help='Số tiền nhỏ hơn trong hai chiều — đây là số sẽ được bù trừ')
    remaining_receivable = fields.Monetary(
        string='Còn lại phải thu', compute='_compute_totals')
    remaining_payable = fields.Monetary(
        string='Còn lại phải trả', compute='_compute_totals')

    @api.depends(
        'receivable_line_ids.selected', 'receivable_line_ids.amount_residual',
        'payable_line_ids.selected', 'payable_line_ids.amount_residual',
    )
    def _compute_totals(self):
        for wiz in self:
            sel_rec = wiz.receivable_line_ids.filtered('selected')
            sel_pay = wiz.payable_line_ids.filtered('selected')
            total_rec = sum(abs(l.amount_residual) for l in sel_rec)
            total_pay = sum(abs(l.amount_residual) for l in sel_pay)
            offset = min(total_rec, total_pay)
            wiz.total_receivable_selected = total_rec
            wiz.total_payable_selected = total_pay
            wiz.amount_to_offset = offset
            wiz.remaining_receivable = total_rec - offset
            wiz.remaining_payable = total_pay - offset

    @api.onchange('partner_id')
    def _onchange_load_lines(self):
        """Tải dòng công nợ chưa đối trừ của đối tác"""
        # Xóa dòng cũ
        self.env['debt.offset.wizard.line'].search(
            [('wizard_id', '=', self.id)]).unlink()

        if not self.partner_id:
            return

        new_lines = []

        # Phải thu: TK 131, dư Nợ (amount_residual > 0)
        rec_domain = [
            ('partner_id', '=', self.partner_id.id),
            ('account_id.account_type', '=', 'trade_receivable'),
            ('reconciled', '=', False),
            ('amount_residual', '>', 0),
            ('move_id.state', '=', 'posted'),
        ]
        for ml in self.env['account.move.line'].search(rec_domain, order='date asc'):
            new_lines.append((0, 0, {
                'side': 'receivable',
                'move_line_id': ml.id,
                'selected': False,
            }))

        # Phải trả: TK 331, dư Có (amount_residual < 0)
        pay_domain = [
            ('partner_id', '=', self.partner_id.id),
            ('account_id.account_type', '=', 'trade_payable'),
            ('reconciled', '=', False),
            ('amount_residual', '<', 0),
            ('move_id.state', '=', 'posted'),
        ]
        for ml in self.env['account.move.line'].search(pay_domain, order='date asc'):
            new_lines.append((0, 0, {
                'side': 'payable',
                'move_line_id': ml.id,
                'selected': False,
            }))

        # Gán tất cả vào wizard
        self.write({'receivable_line_ids': [(5,)] + [l for l in new_lines if l[2]['side'] == 'receivable'],
                    'payable_line_ids': [(5,)] + [l for l in new_lines if l[2]['side'] == 'payable']})

    def action_offset(self):
        """Tạo bút toán bù trừ Dr331 / Cr131 rồi reconcile"""
        sel_rec = self.receivable_line_ids.filtered('selected')
        sel_pay = self.payable_line_ids.filtered('selected')

        if not sel_rec:
            raise UserError(_('Vui lòng chọn ít nhất 1 dòng công nợ phải thu (TK 131).'))
        if not sel_pay:
            raise UserError(_('Vui lòng chọn ít nhất 1 dòng công nợ phải trả (TK 331).'))

        total_rec = sum(abs(l.amount_residual) for l in sel_rec)
        total_pay = sum(abs(l.amount_residual) for l in sel_pay)
        amount = min(total_rec, total_pay)

        if amount <= 0:
            raise UserError(_('Số tiền bù trừ phải lớn hơn 0.'))

        # Lấy tài khoản 131 và 331
        rec_account = sel_rec[0].move_line_id.account_id
        pay_account = sel_pay[0].move_line_id.account_id

        # Tạo journal entry: Nợ 331 / Có 131
        move = self.env['account.move'].create({
            'move_type': 'entry',
            'journal_id': self.journal_id.id,
            'date': self.date,
            'ref': self.ref,
            'line_ids': [
                (0, 0, {
                    'account_id': pay_account.id,   # Nợ TK 331
                    'partner_id': self.partner_id.id,
                    'debit': amount,
                    'credit': 0,
                    'name': self.ref,
                }),
                (0, 0, {
                    'account_id': rec_account.id,   # Có TK 131
                    'partner_id': self.partner_id.id,
                    'debit': 0,
                    'credit': amount,
                    'name': self.ref,
                }),
            ],
        })
        move.action_post()

        # Reconcile TK 131: dòng Có mới với các dòng Nợ từ hóa đơn bán
        new_131_line = move.line_ids.filtered(
            lambda l: l.account_id == rec_account)
        rec_lines = sel_rec.mapped('move_line_id') | new_131_line
        rec_lines.reconcile()

        # Reconcile TK 331: dòng Nợ mới với các dòng Có từ hóa đơn mua
        new_331_line = move.line_ids.filtered(
            lambda l: l.account_id == pay_account)
        pay_lines = sel_pay.mapped('move_line_id') | new_331_line
        pay_lines.reconcile()

        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': _('Bù trừ thành công'),
                'message': _(
                    'Đã tạo bút toán bù trừ %s\n'
                    'Số tiền bù trừ: %s\n'
                    'Nợ 331 / Có 131'
                ) % (move.name, f'{amount:,.0f}'),
                'type': 'success',
                'sticky': False,
            },
        }
