from odoo import models, fields, api
import io
import base64
import openpyxl
from os import path
from openpyxl.styles import PatternFill
from datetime import timedelta


class FinishFetPlanModuleItemPlanHeaderTable(models.Model):
    _name = 'finishfetplanmodule.itemplanheadertable'

    itemplan_id = fields.One2many('finishfetplanmodule.itemplantable', 'itemplanheader_id')
    name = fields.Char('Item ', required=True)
    wo_srno = fields.Char('W.O/ SNo. ', required=True)
    plan_date = fields.Date('Plan Date ', required=True)

    def rescheduledate(self, data, context=None):
        for record in self.itemplan_id:
            if record:
                record.date = self.plan_date + timedelta(days=record.lag_days)

    def button_excel(self, data, context=None):
        fillGRINDING = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
        fillGOUGING = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
        fillWELDING = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
        src = path.dirname(path.realpath(__file__)) + "/FinishFetplan.xlsx"
        wb = openpyxl.load_workbook(src)
        wb.active = 0
        worksheet = wb.active
        filename = 'FinishFetplan.xlsx'
        reportname = "FinishFetPlan:" + self.name
        item_obj = self.env['finishfetplanmodule.itemplanheadertable']
        item_ids = item_obj.search(['&', ('wo_srno', '=', self.wo_srno), ('plan_date', '=', self.plan_date)])
        col = 4
        setcol1 = worksheet.cell(row=14, column=1)
        setcol1.value = item_ids.name or ''
        setcol1 = worksheet.cell(row=14, column=2)
        setcol1.value = item_ids.wo_srno or ''

        for thisitem in item_ids.itemplan_id:
            if thisitem.date >= self.plan_date:
                datediff = (thisitem.date - self.plan_date)
                col = ((datediff.days + 1) * 3) + 1
                dateStr = str(thisitem.date.day) + "/" + str(thisitem.date.month)
                setdate = worksheet.cell(row=12, column=col)
                setdate.value = dateStr

                colorfill = PatternFill(start_color=thisitem.jobrouting_id.colour,
                                        end_color=thisitem.jobrouting_id.colour, fill_type='solid')
                if thisitem.shift_a > 0:
                    setcol2 = worksheet.cell(row=14, column=col)
                    if setcol2.value:
                        thisitem.error_log_a = 'Conflicts with other plan A'
                    else:
                        setcol2.value = thisitem.shift_a or ''
                        setcol2.fill = colorfill
                        if thisitem.error_log_a == 'Conflicts with other plan A':
                            thisitem.error_log_a = ''
                col = col + 1

                if thisitem.shift_b > 0:
                    setcol3 = worksheet.cell(row=14, column=col)
                    if setcol3.value:
                        # valuefill = setcol3.fill
                        thisitem.error_log_b = 'Conflicts with other plan B'
                    else:
                        setcol3.value = thisitem.shift_b or ''
                        setcol3.fill = colorfill
                        if thisitem.error_log_b == 'Conflicts with other plan B':
                            thisitem.error_log_b = ''
                col = col + 1

                if thisitem.shift_c > 0:
                    setcol4 = worksheet.cell(row=14, column=col)
                    if setcol4.value:
                        thisitem.error_log_c = 'Conflicts with other plan C'
                    else:
                        setcol4.value = thisitem.shift_c
                        setcol4.fill = colorfill
                        if thisitem.error_log_c == 'Conflicts with other plan C':
                            thisitem.error_log_c = ''

        fp = io.BytesIO()
        wb.save(fp)
        out = base64.encodestring(fp.getvalue())
        view_empreport_id = self.env['view.empreport'].create(
            {'name': reportname, 'file_name': filename, 'datas_fname': out})
        return {
            'res_id': view_empreport_id.id,
            'name': 'Spent Report',
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': 'view.empreport',
            'view_id': False,
            'type': 'ir.actions.act_window',
        }


class view_empreport(models.TransientModel):
    _name = 'view.empreport'
    _rec_name = 'datas_fname'
    name = fields.Char('Report Name', size=256)
    file_name = fields.Char('File Name', size=256)
    datas_fname = fields.Binary('Report')
