from xlrd import open_workbook
import xlwt


__all__ =('ExcelTable',)

class ExcelTable(object):
    def __init__(self, columns=[], data=[], format=None, extra=[]):
        self.columns = columns
        self.data=data
        self.extra=[]
        
        
    def add_field(self, field_name, default=None):
        """
        Adds a field to the current dataset.
        Raises FieldExistError if the field already exists
        """
        if field_name in self.columns:
            raise ExcelTable.FieldExistError(field_name)
        
        self.columns.append(field_name)
        for row in self.data:
            row[field_name]=None
    

    def add_row(self, row):
        """
        Adds a row to the current dataset
        Raises UndefinedFieldError if a field in 'row' doesn't exist
        among the defined columns.
        """
        for k, v in row.items():
            if not (k in self.columns):
                raise ExcelTable.UndefinedFieldError(k)
        self.data.append(row)
    

    def save(self, filename, sheetname):
        wb_out = xlwt.Workbook(encoding='utf-8')
        sheet_out = wb_out.add_sheet(sheetname)
        row_index=0
        col_key={}
        for i in range(len(self.columns)):
            col_key[self.columns[i]]=i
            sheet_out.write(row_index, i, self.columns[i])
        row_index +=1
        for row in self.data:
            for k, v in row.items():
                sheet_out.write(row_index, col_key[k], v)
            row_index +=1
            
        wb_out.save(filename)
    

    def load_data(self, filename, sheetname):
        wb = open_workbook(filename, formatting_info=True)
        sheet = wb.sheet_by_name(sheetname)
        
        self.columns = []
        self.data=[]
        ncols = sheet.ncols
        for col_index in range(ncols):
            value = sheet.cell(0, col_index).value
            self.columns.append(value)
        
        nrows = sheet.nrows
        for row_index in range(1, nrows):
            row = {}
            for col_index in range(ncols):
                value = sheet.cell(row_index, col_index).value
                row[self.columns[col_index]]=value

            self.data.append(row)
     

    class FieldExistsError(Exception):
        def __init__(self, field_name):
            self.field_name=field_name
        def __str__(self):
            return "FieldExistError: %s" % (self.field_name)
    
        
    class UndefinedFieldError(Exception):
        def __init__(self, field_name):
            self.field_name=field_name
        def __str__(self):
            return "UndefinedFieldError: %s" % (self.field_name)
        