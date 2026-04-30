from odoo import api, fields, models, _
from odoo.exceptions import UserError


class ReconcilePaymentWizardLine(models.TransientModel):
    """Dòng chứng từ chưa đối trừ — hiển thị trong wizard"""
    _name = 'reconcile.payment.wizard.line'
    _description = 'Dòng đối trừ công nợ'

    wizard_id = fields.Many2one('reconcile.payment.wizard', ondelete='cascade')
    selected = fields.Boolean(string='Chọn', default=False)
    move_line_id = fields.Many2one('account.move.line', string='Dòng kế toán', readonly=True)

    # Thông tin hiển thị
    date = fields.Date(related='move_line_id.date', string='Ngày', readonly=True)
    move_name = fields.Char(related='move_line_id.move_id.name', string='Số CT', readonly=True)
    ref = fields.Char(related='move_line_id.ref', string='Diễn giải', readonly=True)
    debit = fields.Monetary(related='move_line_id.debit', string='Nợ', readonly=True)
    credit = fields.Monetary(related='move_line_id.credit', string='Có', readonly=True)
    amount_residual = fields.Monetary(
        related='move_line_id.amount_residual', string='Còn lại', readonly=True)
    currency_id = fields.Many2one(
        related='move_line_id.company_currency_id', readonly=True)
    move_type = fields.Char(
        related='move_line_id.move_id.move_type', string='Loại CT', readonly=True)

    @api.depends('move_line_id')
    def _compute_doc_type(self):
        type_map = {
            'out_invoice': 'Hóa đơn bán',
            'in_invoice': 'Hóa đơn mua',
            'out_refund': 'Trả lại bán',
            'in_refund': 'Trả lại mua',
            'entry': 'Phiếu thu/chi',
        }
        for line in self:
            mt = line.move_line_id.move_id.move_type
            line.doc_type = type_map.get(mt, mt)

    doc_type = fields.Char(string='Loại', compute='_compute_doc_type')


class ReconcilePaymentWizard(models.TransientModel):
    """Wizard ghép phiếu thu/chi với hóa đơn (đối trừ công nợ 131/331)"""
    _name = 'reconcile.payment.wizard'
    _description = 'Ghép phiếu thu/chi với hóa đơn'

    partner_id = fields.Many2one(
        'res.partner', string='Đối tác', required=True)
    account_type = fields.Selection([
        ('trade_receivable', 'Phải thu khách hàng (TK 131)'),
        ('trade_payable',   'Phải trả nhà cung cấp (TK 331)'),
    ], string='Loại công nợ', default='trade_receivable', required=True)

    line_ids = fields.One2many(
        'reconcile.payment.wizard.line', 'wizard_id', string='Danh sách chứng từ')

    # Tổng hợp
    total_debit = fields.Monetary(
        string='Tổng Nợ chọn', compute='_compute_totals')
    total_credit = fields.Monetary(
        string='Tổng Có chọn', compute='_compute_totals')
    total_residual = fields.Monetary(
        string='Tổng còn lại chọn', compute='_compute_totals')
    currency_id = fields.Many2one(
        'res.currency',
        default=lambda self: self.env.company.currency_id)
    count_selected = fields.Integer(
        string='Số dòng chọn', compute='_compute_totals')

    @api.depends('line_ids.selected', 'line_ids.debit', 'line_ids.credit', 'line_ids.amount_residual')
    def _compute_totals(self):
        for wiz in self:
            sel = wiz.line_ids.filtered('selected')
            wiz.total_debit = sum(sel.mapped('debit'))
            wiz.total_credit = sum(sel.mapped('credit'))
            wiz.total_residual = sum(sel.mapped('amount_residual'))
            wiz.count_selected = len(sel)

    @api.onchange('partner_id', 'account_type')
    def _onchange_load_lines(self):
        """Tải danh sách dòng chưa đối trừ khi chọn đối tác / loại công nợ"""
        self.line_ids = [(5, 0, 0)]
        if not self.partner_id or not self.account_type:
            return

        domain = [
            ('partner_id', '=', self.partner_id.id),
            ('account_id.account_type', '=', self.account_type),
            ('reconciled', '=', False),
            ('amount_residual', '!=', 0),
            ('move_id.state', '=', 'posted'),
        ]
        move_lines = self.env['account.move.line'].search(
            domain, order='date asc')

        new_lines = []
        for ml in move_lines:
            new_lines.append((0, 0, {
                'move_line_id': ml.id,
                'selected': False,
            }))
        self.line_ids = new_lines

    def action_select_all(self):
        self.line_ids.write({'selected': True})
        return {
            'type': 'ir.actions.act_window',
            'res_model': self._name,
            'res_id': self.id,
            'view_mode': 'form',
            'target': 'new',
        }

    def action_deselect_all(self):
        self.line_ids.write({'selected': False})
        return {
            'type': 'ir.actions.act_window',
            'res_model': self._name,
            'res_id': self.id,
            'view_mode': 'form',
            'target': 'new',
        }

    def action_reconcile(self):
        """Thực hiện đối trừ các dòng được chọn"""
        selected = self.line_ids.filtered('selected')
        if len(selected) < 2:
            raise UserError(_('Vui lòng chọn ít nhất 2 dòng để đối trừ.'))

        lines = selected.mapped('move_line_id')

        # Kiểm tra chiều (phải có cả Nợ và Có)
        has_debit = any(l.debit > 0 for l in lines)
        has_credit = any(l.credit > 0 for l in lines)
        if not (has_debit and has_credit):
            raise UserError(_(
                'Các dòng chọn phải có cả chiều Nợ và chiều Có để đối trừ nhau.\n'
                'Hóa đơn bán → dòng Nợ 131\n'
                'Phiếu thu    → dòng Có 131'
            ))

        lines.reconcile()

        return {
            'type': 'ir.actions.client',
            'tag': 'display_notification',
            'params': {
                'title': _('Đối trừ thành công'),
                'message': _(
                    'Đã đối trừ %d dòng công nợ. '
                    'Vui lòng kiểm tra lại số dư trên sổ cái.'
                ) % len(selected),
                'type': 'success',
                'sticky': False,
            },
        }
