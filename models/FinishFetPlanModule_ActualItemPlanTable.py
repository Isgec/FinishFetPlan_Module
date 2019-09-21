from odoo import api, fields, models, SUPERUSER_ID, _
from odoo.exceptions import UserError, ValidationError
from reportlab.lib.validators import isNumber


class FinishFetPlanModuleActualItemPlanTable(models.Model):
    _name = 'finishfetplanmodule.actualitemplantable'

    itemplanheader_id = fields.Many2one('finishfetplanmodule.itemplanheadertable', 'actualitemplan_id',
                                        ondelete='cascade')
    name = fields.Char('Item')
    date = fields.Date('Date ', required=True)
    shift_a = fields.Float('A', compute='default_shift_a', store=True, default=0.0)
    shift_b = fields.Float('B', compute='default_shift_b', store=True, default=0.0)
    shift_c = fields.Float('C', compute='default_shift_c', store=True, default=0.0)
    shift_a_c = fields.Char('A')
    shift_b_c = fields.Char('B')
    shift_c_c = fields.Char('C')
    jobrouting_id = fields.Many2one('finishfetplanmodule.jobroutingtable', string='Job Routing')
    error_log_a = fields.Char('Error Log A', readonly=True)
    error_log_b = fields.Char('Error Log B', readonly=True)
    error_log_c = fields.Char('Error Log C', readonly=True)
    lag_days = fields.Integer('Lag Days', compute='default_date', store=True)

    @api.depends('date')
    def default_date(self):
        for record in self:
            if record.date:
                datediff = record.date - record.itemplanheader_id.plan_date
                record.lag_days = datediff.days
                record.name = record.itemplanheader_id.name

    @api.depends('shift_a_c')
    def default_shift_a(self):
        for record in self:
            if isNumber(record.shift_a_c):
                record.shift_a = record.shift_a_c

    @api.depends('shift_b_c')
    def default_shift_b(self):
        for record in self:
            if isNumber(record.shift_b_c):
                record.shift_b = record.shift_b_c

    @api.depends('shift_c_c')
    def default_shift_c(self):
        for record in self:
            if isNumber(record.shift_c_c):
                record.shift_c = record.shift_c_c
