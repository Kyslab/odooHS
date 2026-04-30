{
    'name': 'Quản lý Tiền mặt (VN)',
    'version': '17.0.1.0.0',
    'summary': 'Giao diện quản lý thu chi tiền mặt & ngân hàng kiểu MISA',
    'description': '''
        Module quản lý quỹ tiền mặt & ngân hàng:
        - Danh sách phiếu thu/chi tiền mặt (TK 111x)
        - Danh sách phiếu thu/chi ngân hàng (TK 112x)
        - Tạo phiếu thu/chi nhanh (tiền mặt & ngân hàng)
        - Thu/Trả tiền theo hóa đơn (tiền mặt & ngân hàng)
        - Lọc theo ngày, đối tượng, lý do, diễn giải
    ''',
    'author': 'Custom VN',
    'category': 'Accounting/Vietnam',
    'depends': ['account'],
    'data': [
        'security/ir.model.access.csv',
        # ── Tiền mặt ──
        'wizards/cash_receipt_wizard_views.xml',
        'wizards/cash_payment_wizard_views.xml',
        'wizards/cash_receipt_invoice_wizard_views.xml',
        'wizards/cash_payment_invoice_wizard_views.xml',
        'views/cash_move_views.xml',
        # ── Ngân hàng ──
        'wizards/bank_receipt_wizard_views.xml',
        'wizards/bank_payment_wizard_views.xml',
        'wizards/bank_receipt_invoice_wizard_views.xml',
        'wizards/bank_payment_invoice_wizard_views.xml',
        'views/bank_move_views.xml',
        # ── Đối trừ công nợ ──
        'wizards/reconcile_payment_wizard_views.xml',
        'wizards/debt_offset_wizard_views.xml',
        # ── Menu ──
        'views/cash_menu.xml',
    ],
    'assets': {
        'web.assets_backend': [
            'l10n_vn_cash_manager/static/src/css/cash_style.css',
        ],
    },
    'installable': True,
    'application': False,
    'license': 'LGPL-3',
}
