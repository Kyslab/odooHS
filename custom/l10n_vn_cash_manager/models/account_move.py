from odoo import models, fields, api


class AccountMove(models.Model):
    _inherit = 'account.move'

    # Lý do thu/chi — người dùng có thể nhập hoặc để tự động
    x_ly_do = fields.Char(
        string='Lý do thu/chi',
        help='Lý do thu tiền hoặc chi tiền (nhập thủ công hoặc tự động)'
    )

    # Loại chứng từ tiền mặt (tính toán)
    x_loai_ct = fields.Selection([
        ('thu', 'Thu tiền'),
        ('chi', 'Chi tiền'),
    ], string='Loại', compute='_compute_loai_ct', store=True)

    # Số tiền thực tế trên TK quỹ (111x/112x) — KHÔNG cộng TK chiết khấu hay TK khác
    x_so_tien_quy = fields.Monetary(
        string='Số tiền',
        currency_field='currency_id',
        compute='_compute_so_tien_quy',
        store=True,
    )

    # Diễn giải tổng hợp từ các dòng
    x_dien_giai = fields.Char(
        string='Diễn giải',
        compute='_compute_dien_giai',
        store=True
    )

    @api.depends('line_ids', 'line_ids.debit', 'line_ids.credit',
                 'line_ids.account_id', 'journal_id')
    def _compute_loai_ct(self):
        """Xác định Thu hay Chi dựa vào TK tiền mặt/ngân hàng"""
        for move in self:
            loai = False
            if move.journal_id and move.journal_id.type in ('cash', 'bank'):
                # Tìm dòng TK tiền (111x, 112x)
                for line in move.line_ids:
                    code = line.account_id.code or ''
                    if code.startswith(('111', '112')):
                        if line.debit > 0:
                            loai = 'thu'   # Tiền vào quỹ = Thu tiền
                        elif line.credit > 0:
                            loai = 'chi'   # Tiền ra quỹ = Chi tiền
                        break
            move.x_loai_ct = loai

    @api.depends('line_ids', 'line_ids.debit', 'line_ids.credit',
                 'line_ids.account_id', 'journal_id')
    def _compute_so_tien_quy(self):
        """Lấy đúng số tiền trên dòng TK quỹ (111x/112x) — không cộng các TK khác"""
        for move in self:
            so_tien = 0.0
            if move.journal_id and move.journal_id.type in ('cash', 'bank'):
                for line in move.line_ids:
                    code = line.account_id.code or ''
                    if code.startswith(('111', '112')):
                        # Thu: Nợ TK tiền; Chi: Có TK tiền
                        so_tien = line.debit if line.debit > 0 else line.credit
                        break
            move.x_so_tien_quy = so_tien

    @api.depends('line_ids', 'line_ids.name', 'narration', 'ref')
    def _compute_dien_giai(self):
        """Lấy diễn giải từ dòng đầu tiên có nội dung"""
        for move in self:
            dien_giai = move.narration or move.ref or ''
            if not dien_giai:
                # Lấy từ dòng hạch toán đầu tiên (bỏ qua dòng trống)
                for line in move.line_ids:
                    if line.name and line.name != '/':
                        dien_giai = line.name
                        break
            move.x_dien_giai = dien_giai or move.name or ''
