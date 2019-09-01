from odoo import models, fields


class FinishFetPlanModuleManpowerTable(models.Model):
    _name = 'finishfetplanmodule.manpowertable'

    name = fields.Char('Name', required=True)
    shift_a = fields.Integer('Shift A', required=True)
    shift_b = fields.Integer('Shift B', required=True)
    shift_c = fields.Integer('Shift C', required=True)
    jobrouting_id = fields.Many2one('finishfetplanmodule.jobroutingtable', string='Job Routing', required=True)
