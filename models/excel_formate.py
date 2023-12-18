from odoo import fields,models

class ExcelFormateList(models.Model):
    _name = 'excel.formate.list'

    name = fields.Char()
    list_length = fields.Integer()
    field_lines = fields.One2many('excel.formate.list.lines','excel_id')

class ExcelFormateListLines(models.Model):
    _name = 'excel.formate.list.lines'

    excel_id = fields.Many2one('excel.formate.list')
    field_id = fields.Char()
    excel_name = fields.Char()
    manyonename = fields.Char("Many2One Name",default="name")
    field_type = fields.Selection([('normal','Normal'),('many2one','Many2One'),('date','Date')],required=1,default='normal')

