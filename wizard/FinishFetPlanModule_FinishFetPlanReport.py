from odoo import models, fields
import io
import base64
import openpyxl
from os import path
from openpyxl.styles import PatternFill


class FinishFetPlanModule_FinishFetPlanReport(models.TransientModel):
    _name = 'finishfetplanmodule.finishfetplanreport'
    from_dt = fields.Date('From Date ', required=True)
    name = fields.Char('Report Name')

    def button_excel(self, data, context=None):
        fillGRINDING = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
        fillGOUGING = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
        fillWELDING = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')
        src = path.dirname(path.realpath(__file__)) + "/FinishFetplan.xlsx"
        wb = openpyxl.load_workbook(src)
        wb.active = 0
        worksheet = wb.active
        filename = 'FinishFetplan.xlsx'
        reportname = "FinishFetPlan:"
        self.name = 'Finish Fet Plan Report as on: ' + str(self.from_dt)

        itemheader_obj = self.env['finishfetplanmodule.itemplanheadertable']
        itemheader_ids = itemheader_obj.search([(1, '=', 1)])
        row = 14
        for thisitems_id in itemheader_ids:
            setcol1 = worksheet.cell(row=row, column=1)
            setcol1.value = thisitems_id.name or ''
            setcol1 = worksheet.cell(row=row, column=2)
            setcol1.value = thisitems_id.wo_srno or ''

            for thisitem in thisitems_id.itemplan_id:
                if thisitem.date >= self.from_dt:
                    datediff = (thisitem.date - self.from_dt)
                    col = ((datediff.days + 1) * 3) + 1
                    dateStr = str(thisitem.date.day) + "/" + str(thisitem.date.month)
                    setdate = worksheet.cell(row=12, column=col)
                    setdate.value = dateStr

                    colorfill = PatternFill(start_color=thisitem.jobrouting_id.colour,
                                            end_color=thisitem.jobrouting_id.colour, fill_type='solid')
                    if thisitem.shift_a > 0:
                        setcol2 = worksheet.cell(row=row, column=col)
                        if setcol2.value:
                            thisitem.error_log_a = 'Conflicts with other plan A'
                        else:
                            setcol2.value = thisitem.shift_a or ''
                            setcol2.fill = colorfill
                            if thisitem.error_log_a == 'Conflicts with other plan A':
                                thisitem.error_log_a = ''
                    col = col + 1

                    if thisitem.shift_b > 0:
                        setcol3 = worksheet.cell(row=row, column=col)
                        if setcol3.value:
                            thisitem.error_log_b = 'Conflicts with other plan B'
                        else:
                            setcol3.value = thisitem.shift_b or ''
                            setcol3.fill = colorfill
                            if thisitem.error_log_b == 'Conflicts with other plan B':
                                thisitem.error_log_b = ''
                    col = col + 1

                    if thisitem.shift_c > 0:
                        setcol4 = worksheet.cell(row=row, column=col)
                        if setcol4.value:
                            thisitem.error_log_c = 'Conflicts with other plan C'
                        else:
                            setcol4.value = thisitem.shift_c
                            setcol4.fill = colorfill
                            if thisitem.error_log_c == 'Conflicts with other plan C':
                                thisitem.error_log_c = ''
            row = row + 2

        wb.active = 1
        worksheet = wb.active
        itemheader_obj = self.env['finishfetplanmodule.manpowertable']
        item_ids = itemheader_obj.search([(1, '=', 1)])
        for thisitem in item_ids:
            if thisitem.jobrouting_id.name == 'WELDING':
                welding_shift_a = thisitem.shift_a
                welding_shift_b = thisitem.shift_b
                welding_shift_c = thisitem.shift_c
            if thisitem.jobrouting_id.name == 'GRINDING':
                grinding_shift_a = thisitem.shift_a
                grinding_shift_b = thisitem.shift_b
                grinding_shift_c = thisitem.shift_c
            if thisitem.jobrouting_id.name == 'Gouging':
                gouging_shift_a = thisitem.shift_a
                gouging_shift_b = thisitem.shift_b
                gouging_shift_c = thisitem.shift_c

        itemheader_obj = self.env['finishfetplanmodule.itemplantable']
        item_ids = itemheader_obj.search([('date', '>=', self.from_dt)])

        for thisitem in item_ids:
            if thisitem.date >= self.from_dt:
                datediff = (thisitem.date - self.from_dt)
                col = ((datediff.days + 1) * 3)
                dateStr = str(thisitem.date.day) + "/" + str(thisitem.date.month)
                setdate = worksheet.cell(row=2, column=col)
                setdate.value = dateStr
                if thisitem.jobrouting_id.name == 'WELDING':
                    row = 5
                    if thisitem.shift_a > 0:
                        setcol = worksheet.cell(row=row-1, column=col)
                        setcol.value = welding_shift_a
                        setcol2 = worksheet.cell(row=row, column=col)
                        if setcol2.value:
                            setcol2.value = setcol2.value + thisitem.shift_a or ''
                        else:
                            setcol2.value = thisitem.shift_a or ''
                    col = col + 1

                    if thisitem.shift_b > 0:
                        setcol = worksheet.cell(row=row-1, column=col)
                        setcol.value = welding_shift_b
                        setcol3 = worksheet.cell(row=row, column=col)
                        if setcol3.value:
                            setcol3.value =  setcol3.value + thisitem.shift_b or ''
                        else:
                            setcol3.value = thisitem.shift_b or ''
                    col = col + 1

                    if thisitem.shift_c > 0:
                        setcol = worksheet.cell(row=row-1, column=col)
                        setcol.value = welding_shift_c
                        setcol4 = worksheet.cell(row=row, column=col)
                        if setcol4.value:
                            setcol4.value = setcol4.value + thisitem.shift_c
                        else:
                            setcol4.value = thisitem.shift_c

                if thisitem.jobrouting_id.name == 'GRINDING':
                    row = 8
                    if thisitem.shift_a > 0:
                        setcol = worksheet.cell(row=row-1, column=col)
                        setcol.value = grinding_shift_a
                        setcol2 = worksheet.cell(row=row, column=col)
                        if setcol2.value:
                            setcol2.value = setcol2.value + thisitem.shift_a or ''
                        else:
                            setcol2.value = thisitem.shift_a or ''
                    col = col + 1

                    if thisitem.shift_b > 0:
                        setcol = worksheet.cell(row=row-1, column=col)
                        setcol.value = grinding_shift_b
                        setcol3 = worksheet.cell(row=row, column=col)
                        if setcol3.value:
                            setcol3.value =  setcol3.value + thisitem.shift_b or ''
                        else:
                            setcol3.value = thisitem.shift_b or ''
                    col = col + 1

                    if thisitem.shift_c > 0:
                        setcol = worksheet.cell(row=row-1, column=col)
                        setcol.value = grinding_shift_c
                        setcol4 = worksheet.cell(row=row, column=col)
                        if setcol4.value:
                            setcol4.value = setcol4.value + thisitem.shift_c
                        else:
                            setcol4.value = thisitem.shift_c

                if thisitem.jobrouting_id.name == 'Gouging':
                    row = 11
                    if thisitem.shift_a > 0:
                        setcol = worksheet.cell(row=row-1, column=col)
                        setcol.value = gouging_shift_a
                        setcol2 = worksheet.cell(row=row, column=col)
                        if setcol2.value:
                            setcol2.value = setcol2.value + thisitem.shift_a or ''
                        else:
                            setcol2.value = thisitem.shift_a or ''
                    col = col + 1

                    if thisitem.shift_b > 0:
                        setcol = worksheet.cell(row=row-1, column=col)
                        setcol.value = gouging_shift_b
                        setcol3 = worksheet.cell(row=row, column=col)
                        if setcol3.value:
                            setcol3.value = setcol3.value + thisitem.shift_b or ''
                        else:
                            setcol3.value = thisitem.shift_b or ''
                    col = col + 1

                    if thisitem.shift_c > 0:
                        setcol = worksheet.cell(row=row-1, column=col)
                        setcol.value = gouging_shift_c
                        setcol4 = worksheet.cell(row=row, column=col)
                        if setcol4.value:
                            setcol4.value = setcol4.value + thisitem.shift_c
                        else:
                            setcol4.value = thisitem.shift_c

        wb.active = 0
        worksheet = wb.active
        fp = io.BytesIO()
        wb.save(fp)
        out = base64.encodestring(fp.getvalue())
        view_ffpreport_id = self.env['view.ffpreport'].create(
            {'name': reportname, 'file_name': filename, 'datas_fname': out})
        return {
            'res_id': view_ffpreport_id.id,
            'name': 'Finish Fet Plan Report',
            'view_type': 'form',
            'view_mode': 'form',
            'res_model': 'view.ffpreport',
            'view_id': False,
            'type': 'ir.actions.act_window',
        }


class view_ffpreport(models.TransientModel):
    _name = 'view.ffpreport'
    _rec_name = 'datas_fname'
    name = fields.Char('Report Name', size=256)
    file_name = fields.Char('File Name', size=256)
    datas_fname = fields.Binary('Report')
