from odoo import api, fields, models, SUPERUSER_ID, _
from odoo.exceptions import UserError, ValidationError


class FinishFetPlanModuleActualItemPlanTable(models.Model):
    _name = 'finishfetplanmodule.actualitemplantable'

    itemplanheader_id = fields.Many2one('finishfetplanmodule.itemplanheadertable', 'actualitemplan_id', ondelete='cascade')
    name = fields.Char('Item')
    date = fields.Date('Date ', required=True)
    shift_a = fields.Integer('A', required=True)
    shift_b = fields.Integer('B', required=True)
    shift_c = fields.Integer('C', required=True)
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
