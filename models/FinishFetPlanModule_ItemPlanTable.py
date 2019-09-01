from odoo import models, fields, api


class FinishFetPlanModuleItemPlanTable(models.Model):
    _name = 'finishfetplanmodule.itemplantable'

    itemplanheader_id = fields.Many2one('finishfetplanmodule.itemplanheadertable', 'itemplan_id')
    name = fields.Char('Item ', required=True)
    date = fields.Date('Date ', required=True)
    shift_a = fields.Integer('A', required=True)
    shift_b = fields.Integer('B', required=True)
    shift_c = fields.Integer('C', required=True)
    jobrouting_id = fields.Many2one('finishfetplanmodule.jobroutingtable', string='Job Routing', required=True)
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
