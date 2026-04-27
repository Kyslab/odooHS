from odoo import models, fields, api, _
from odoo.exceptions import UserError


class BankReceiptWizard(models.TransientModel):
    """Wizard tạo phiếu thu ngân hàng nhanh"""
    _name = 'bank.receipt.wizard'
    _description = 'Phiếu Thu Ngân Hàng'

    date = fields.Date(
        string='Ngày hạch toán',
        required=True,
        default=fields.Date.context_today
    )
    partner_id = fields.Many2one(
        'res.partner',
        string='Đối tượng (Khách hàng)',
    )
    amount = fields.Monetary(
        string='Số tiền thu',
        required=True,
        currency_field='currency_id'
    )
    currency_id = fields.Many2one(
        'res.currency',
        default=lambda self: self.env.company.currency_id
    )
    account_id = fields.Many2one(
        'account.account',
        string='TK đối ứng (Có)',
        required=True,
        help='Tài khoản ghi Có khi thu tiền (VD: TK131, TK511...)',
        domain=[('reconcile', '=', True)]
    )
    journal_id = fields.Many2one(
        'account.journal',
        string='Tài khoản ngân hàng',
        required=True,
        domain=[('type', '=', 'bank')],
        default=lambda self: self._default_journal()
    )
    ly_do = fields.Char(
        string='Lý do thu tiền',
        default='Phiếu thu ngân hàng'
    )
    dien_giai = fields.Char(
        string='Diễn giải',
    )

    # ── Phân tích ──────────────────────────────────────────────────────────
    ma_thong_ke_id = fields.Many2one(
        'account.analytic.account',
        string='Mã thống kê',
        domain="[('plan_id.name', 'ilike', 'thống kê')]"
    )
    khoan_muc_id = fields.Many2one(
        'account.analytic.account',
        string='Khoản mục CP',
        domain="[('plan_id.name', 'ilike', 'khoản mục')]"
    )
    cong_trinh_id = fields.Many2one(
        'account.analytic.account',
        string='Tên công trình',
        domain="[('plan_id.name', 'in', ['Projects', 'Dự án'])]"
    )

    def _default_journal(self):
        return self.env['account.journal'].search(
            [('type', '=', 'bank'), ('company_id', '=', self.env.company.id)],
            limit=1
        )

    def _build_analytic_distribution(self):
        dist = {}
        if self.ma_thong_ke_id:
            dist[str(self.ma_thong_ke_id.id)] = 100
        if self.khoan_muc_id:
            dist[str(self.khoan_muc_id.id)] = 100
        if self.cong_trinh_id:
            dist[str(self.cong_trinh_id.id)] = 100
        return dist or False

    def action_tao_phieu_thu(self):
        """Tạo bút toán phiếu thu ngân hàng"""
        self.ensure_one()
        if self.amount <= 0:
            raise UserError(_('Số tiền phải lớn hơn 0!'))

        bank_account = self.journal_id.default_account_id
        if not bank_account:
            raise UserError(_('Tài khoản ngân hàng chưa có tài khoản mặc định. '
                              'Vào Cấu hình → Nhật ký để thiết lập.'))

        dien_giai = self.dien_giai or self.ly_do or 'Phiếu thu ngân hàng'
        analytic_dist = self._build_analytic_distribution()

        # Dòng TK đối ứng (Có) — gắn phân tích nếu có
        doi_ung_line = {
            'account_id': self.account_id.id,
            'partner_id': self.partner_id.id if self.partner_id else False,
            'debit': 0,
            'credit': self.amount,
            'name': dien_giai,
        }
        if analytic_dist:
            doi_ung_line['analytic_distribution'] = analytic_dist

        move_vals = {
            'move_type': 'entry',
            'date': self.date,
            'journal_id': self.journal_id.id,
            'partner_id': self.partner_id.id if self.partner_id else False,
            'ref': dien_giai,
            'x_ly_do': self.ly_do,
            'line_ids': [
                (0, 0, {
                    'account_id': bank_account.id,
                    'partner_id': self.partner_id.id if self.partner_id else False,
                    'debit': self.amount,
                    'credit': 0,
                    'name': dien_giai,
                }),
                (0, 0, doi_ung_line),
            ]
        }

        move = self.env['account.move'].create(move_vals)
        move.action_post()

        return {
            'type': 'ir.actions.act_window',
            'name': _('Phiếu Thu Ngân Hàng'),
            'res_model': 'account.move',
            'res_id': move.id,
            'view_mode': 'form',
            'target': 'current',
        }
