from odoo import models, fields
import io
import base64
import openpyxl
from datetime import timedelta
from os import path
from openpyxl.styles import PatternFill
from tempfile import TemporaryFile


class FinishFetPlanModule_FinishFetPlanReport(models.TransientModel):
    _name = 'finishfetplanmodule.finishfetplanreport'
    from_dt = fields.Date('From Date ', required=True)
    name = fields.Char('Report Name')
    upload_file = fields.Binary(string="Upload File")
    uploadedfilename = fields.Char('File Name', size=256, default='Chose File')
    readfromexcel = fields.Text('Sample Read')

    def upload_excel(self, data, context=None):
        # Generating of the excel file to be read by openpyxl
        file = base64.decodestring(self.upload_file)
        excel_fileobj = TemporaryFile('wb+')
        excel_fileobj.write(file)
        excel_fileobj.seek(0)

        # Create workbook
        wb = openpyxl.load_workbook(excel_fileobj, data_only=True)
        wb.active = 0
        worksheet = wb.active

        header_obj = self.env['finishfetplanmodule.itemplanheadertable']
        header_ids = header_obj.search([(1, '=', 1)])
        self.readfromexcel = 'Start:'
        itempos = 14
        relativedate = 0
        my_max_col = 100
        for thisheader_ids in header_ids:
            self.readfromexcel = self.readfromexcel + '{ Items : ' + thisheader_ids.name + '}'
            relativedate = 0
            for thisitems_ids in thisheader_ids.itemplan_id:
                if thisitems_ids.date >= self.from_dt:
                    jobroutingid = thisitems_ids.jobrouting_id.id
                    plandate = thisitems_ids.date
                    self.readfromexcel = self.readfromexcel + ', Unlinked->' + str(thisitems_ids.date)
                    thisitems_ids.unlink()  # Delete Record  from Item table With particular

            for i in range(4, my_max_col + 1, 3):
                # Reading for Shift A
                getdt = self.from_dt + timedelta(relativedate)
                getcol = worksheet.cell(row=itempos, column=i)
                jobrouting_obj = self.env['finishfetplanmodule.jobroutingtable']
                jobrouting_id = jobrouting_obj.search([('colour', '=', str(getcol.fill)[139:147])])

                for thisjob in jobrouting_id:
                    self.readfromexcel = self.readfromexcel + \
                                         ' {Job:' + thisjob.name + \
                                         ' Date:' + str(getdt) + \
                                         'Shift A:' + str(getcol.value) + \
                                         '} '
                    # add record  for Shift A
                    thisheader_ids.itemplan_id.create(
                        {'itemplanheader_id': thisheader_ids.id, 'jobrouting_id': thisjob.id,
                         'date': getdt,
                         'name': 'Added Shift A',
                         'shift_a': getcol.value,
                         'shift_b': 0,
                         'shift_c': 0})

                # Reading for Shift B
                getdt = self.from_dt + timedelta(relativedate)
                getcol = worksheet.cell(row=itempos, column=i + 1)
                jobrouting_obj = self.env['finishfetplanmodule.jobroutingtable']
                jobrouting_id = jobrouting_obj.search([('colour', '=', str(getcol.fill)[139:147])])

                for thisjob in jobrouting_id:
                    self.readfromexcel = self.readfromexcel + \
                                         ' {Job:' + thisjob.name + \
                                         ' Date:' + str(getdt) + \
                                         'Shift B:' + str(getcol.value) + \
                                         '} '
                    # add record  for Shift B
                    thisheader_ids.itemplan_id.create(
                        {'itemplanheader_id': thisheader_ids.id, 'jobrouting_id': thisjob.id,
                         'date': getdt,
                         'name': 'Added Shift B',
                         'shift_a': 0,
                         'shift_b': getcol.value,
                         'shift_c': 0})

                # Reading for Shift C
                getdt = self.from_dt + timedelta(relativedate)
                getcol = worksheet.cell(row=itempos, column=i + 2)
                jobrouting_obj = self.env['finishfetplanmodule.jobroutingtable']
                jobrouting_id = jobrouting_obj.search([('colour', '=', str(getcol.fill)[139:147])])

                for thisjob in jobrouting_id:
                    self.readfromexcel = self.readfromexcel + \
                                         ' {Job:' + thisjob.name + \
                                         ' Date:' + str(getdt) + \
                                         'Shift C:' + str(getcol.value) + \
                                         '} '
                    # add record  for Shift C
                    thisheader_ids.itemplan_id.create(
                        {'itemplanheader_id': thisheader_ids.id, 'jobrouting_id': thisjob.id,
                         'date': getdt,
                         'name': 'Added Shift C',
                         'shift_a': 0,
                         'shift_b': 0,
                         'shift_c': getcol.value})

                relativedate = relativedate + 1
            itempos = itempos + 2

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
                        setcol = worksheet.cell(row=row - 1, column=col)
                        setcol.value = welding_shift_a
                        setcol2 = worksheet.cell(row=row, column=col)
                        if setcol2.value:
                            setcol2.value = setcol2.value + thisitem.shift_a or ''
                        else:
                            setcol2.value = thisitem.shift_a or ''
                    col = col + 1

                    if thisitem.shift_b > 0:
                        setcol = worksheet.cell(row=row - 1, column=col)
                        setcol.value = welding_shift_b
                        setcol3 = worksheet.cell(row=row, column=col)
                        if setcol3.value:
                            setcol3.value = setcol3.value + thisitem.shift_b or ''
                        else:
                            setcol3.value = thisitem.shift_b or ''
                    col = col + 1

                    if thisitem.shift_c > 0:
                        setcol = worksheet.cell(row=row - 1, column=col)
                        setcol.value = welding_shift_c
                        setcol4 = worksheet.cell(row=row, column=col)
                        if setcol4.value:
                            setcol4.value = setcol4.value + thisitem.shift_c
                        else:
                            setcol4.value = thisitem.shift_c

                if thisitem.jobrouting_id.name == 'GRINDING':
                    row = 8
                    if thisitem.shift_a > 0:
                        setcol = worksheet.cell(row=row - 1, column=col)
                        setcol.value = grinding_shift_a
                        setcol2 = worksheet.cell(row=row, column=col)
                        if setcol2.value:
                            setcol2.value = setcol2.value + thisitem.shift_a or ''
                        else:
                            setcol2.value = thisitem.shift_a or ''
                    col = col + 1

                    if thisitem.shift_b > 0:
                        setcol = worksheet.cell(row=row - 1, column=col)
                        setcol.value = grinding_shift_b
                        setcol3 = worksheet.cell(row=row, column=col)
                        if setcol3.value:
                            setcol3.value = setcol3.value + thisitem.shift_b or ''
                        else:
                            setcol3.value = thisitem.shift_b or ''
                    col = col + 1

                    if thisitem.shift_c > 0:
                        setcol = worksheet.cell(row=row - 1, column=col)
                        setcol.value = grinding_shift_c
                        setcol4 = worksheet.cell(row=row, column=col)
                        if setcol4.value:
                            setcol4.value = setcol4.value + thisitem.shift_c
                        else:
                            setcol4.value = thisitem.shift_c

                if thisitem.jobrouting_id.name == 'Gouging':
                    row = 11
                    if thisitem.shift_a > 0:
                        setcol = worksheet.cell(row=row - 1, column=col)
                        setcol.value = gouging_shift_a
                        setcol2 = worksheet.cell(row=row, column=col)
                        if setcol2.value:
                            setcol2.value = setcol2.value + thisitem.shift_a or ''
                        else:
                            setcol2.value = thisitem.shift_a or ''
                    col = col + 1

                    if thisitem.shift_b > 0:
                        setcol = worksheet.cell(row=row - 1, column=col)
                        setcol.value = gouging_shift_b
                        setcol3 = worksheet.cell(row=row, column=col)
                        if setcol3.value:
                            setcol3.value = setcol3.value + thisitem.shift_b or ''
                        else:
                            setcol3.value = thisitem.shift_b or ''
                    col = col + 1

                    if thisitem.shift_c > 0:
                        setcol = worksheet.cell(row=row - 1, column=col)
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
