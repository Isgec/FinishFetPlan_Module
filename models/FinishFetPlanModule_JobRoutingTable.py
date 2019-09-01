from odoo import models, fields


class FinishFetPlanModuleJobRoutingTable(models.Model):
    _name = 'finishfetplanmodule.jobroutingtable'
    name = fields.Char('Name', required=True)
    sequence_no = fields.Integer('Sequence', required=True)
    colour = fields.Char('CellColourCode', required=True)
