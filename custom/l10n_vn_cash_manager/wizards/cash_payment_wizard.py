from odoo import models, fields, api, _
from odoo.exceptions import UserError


class CashPaymentWizard(models.TransientModel):
    """Wizard tạo phiếu chi tiền mặt nhanh"""
    _name = 'cash.payment.wizard'
    _description = 'Phiếu Chi Tiền Mặt'

    date = fields.Date(
        string='Ngày hạch toán',
        required=True,
        default=fields.Date.context_today
    )
    partner_id = fields.Many2one(
        'res.partner',
        string='Đối tượng',
    )
    amount = fields.Monetary(
        string='Số tiền chi',
        required=True,
        currency_field='currency_id'
    )
    currency_id = fields.Many2one(
        'res.currency',
        default=lambda self: self.env.company.currency_id
    )
    account_id = fields.Many2one(
        'account.account',
        string='TK đối ứng (Nợ)',
        required=True,
        help='Tài khoản ghi Nợ khi chi tiền (VD: TK331 - Trả NCC, TK641 - Chi phí...)',
    )
    journal_id = fields.Many2one(
        'account.journal',
        string='Quỹ tiền mặt',
        required=True,
        domain=[('type', '=', 'cash')],
        default=lambda self: self._default_journal()
    )
    ly_do = fields.Char(
        string='Lý do chi tiền',
        default='Phiếu chi tiền mặt'
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
        journal = self.env['account.journal'].search(
            [('type', '=', 'cash'), ('company_id', '=', self.env.company.id)],
            limit=1
        )
        return journal

    def _build_analytic_distribution(self):
        dist = {}
        if self.ma_thong_ke_id:
            dist[str(self.ma_thong_ke_id.id)] = 100
        if self.khoan_muc_id:
            dist[str(self.khoan_muc_id.id)] = 100
        if self.cong_trinh_id:
            dist[str(self.cong_trinh_id.id)] = 100
        return dist or False

    def action_tao_phieu_chi(self):
        """Tạo bút toán phiếu chi tiền mặt"""
        self.ensure_one()
        if self.amount <= 0:
            raise UserError(_('Số tiền phải lớn hơn 0!'))

        cash_account = self.journal_id.default_account_id
        if not cash_account:
            raise UserError(_('Quỹ tiền mặt chưa có tài khoản mặc định. '
                              'Vào Cấu hình → Nhật ký để thiết lập.'))

        dien_giai = self.dien_giai or self.ly_do or 'Phiếu chi tiền mặt'
        analytic_dist = self._build_analytic_distribution()

        # Dòng TK đối ứng (Nợ) — gắn phân tích nếu có
        doi_ung_line = {
            'account_id': self.account_id.id,
            'partner_id': self.partner_id.id if self.partner_id else False,
            'debit': self.amount,
            'credit': 0,
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
                (0, 0, doi_ung_line),
                # Dòng Có: TK tiền mặt giảm (tiền ra khỏi quỹ)
                (0, 0, {
                    'account_id': cash_account.id,
                    'partner_id': self.partner_id.id if self.partner_id else False,
                    'debit': 0,
                    'credit': self.amount,
                    'name': dien_giai,
                }),
            ]
        }

        move = self.env['account.move'].create(move_vals)
        move.action_post()

        return {
            'type': 'ir.actions.act_window',
            'name': _('Phiếu Chi'),
            'res_model': 'account.move',
            'res_id': move.id,
            'view_mode': 'form',
            'target': 'current',
        }
