from odoo import fields, models
from odoo.exceptions import UserError


class PartnerXlsx(models.AbstractModel):
    _name = "report.enz_excel_template.contactxlsx"
    _inherit = "report.report_xlsx.abstract"
    _description = "Partner XLSX Report"

    def generate_xlsx_report(self, workbook, data, partners):
        sheet = workbook.add_worksheet("Report")
        bold = workbook.add_format({"bold": True})
        date = workbook.add_format({'text_wrap': True, 'num_format': 'yyyy-mm-dd'})
        contact_field = self.env['excel.formate.list'].search([('name', '=', 'enz_excel_template.contactxlsx')])
        label_list = ['Name', 'Street', 'Street2', 'City', 'State',
                      'Country', 'Zip', 'Tax ID', 'Phone', 'Mobile',
                      'Email', 'Website Link', 'Customer Rank', 'Supplier Rank']
        length = 14
        if contact_field:
            length = length + len(contact_field.field_lines)
            label_list = label_list + contact_field.field_lines.mapped('excel_name')
        for line in range(0, length):
            sheet.write(0, line, label_list[line], bold)
        i = 1
        for obj in partners:
            # if obj.name:
            #     sheet.write(i, 0, getattr(obj, "name"))
            # if obj.street:
            #     sheet.write(i, 1, obj.street)
            if obj.name:
                sheet.write(i, 0, obj.name)
            if obj.street:
                sheet.write(i, 1, obj.street)
            if obj.street2:
                sheet.write(i, 2, obj.street2)
            if obj.city:
                sheet.write(i, 3, obj.city)
            if obj.state_id:
                sheet.write(i, 4, obj.state_id.name)
            if obj.country_id:
                sheet.write(i, 5, obj.country_id.name)
            if obj.zip:
                sheet.write(i, 6, obj.zip)
            if obj.vat:
                sheet.write(i, 7, obj.vat)
            if obj.phone:
                sheet.write(i, 8, obj.phone)
            if obj.mobile:
                sheet.write(i, 9, obj.mobile)
            if obj.email:
                sheet.write(i, 10, obj.email)
            if obj.website:
                sheet.write(i, 11, obj.website)
            if obj.customer_rank:
                sheet.write(i, 12, obj.customer_rank)
            if obj.supplier_rank:
                sheet.write(i, 13, obj.supplier_rank)
            ##CALLING START
            total_list_no = 14
            if contact_field:
                if contact_field.field_lines:
                    for listline in contact_field.field_lines:
                        model_id = self.env['ir.model'].sudo().search([('model', '=', 'res.partner')]).id
                        field_id = self.env['ir.model.fields'].sudo().search(
                            [('name', '=', listline.field_id), ('model_id', '=', model_id)])
                        if field_id:
                            if getattr(obj, listline.field_id):
                                if listline.field_type == 'normal':
                                    sheet.write(i, total_list_no, getattr(obj, listline.field_id))
                                elif listline.field_type == 'many2one':
                                    manyone = getattr(obj, listline.field_id)
                                    if manyone:
                                        sheet.write(i, total_list_no, getattr(manyone, listline.manyonename))
                                else:
                                    sheet.write(i, total_list_no, getattr(obj, listline.field_id), date)
                            total_list_no = total_list_no + 1
                        else:
                            raise UserError("No Field Found in the Name " + listline.field_id)
            ##CALLING END
            i += 1
