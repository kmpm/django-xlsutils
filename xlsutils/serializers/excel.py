from django.db import models
from django.utils.encoding import smart_unicode, is_protected_type
from django.core.serializers import base
from django.db import transaction
import xlrd
import xlwt
from datetime import datetime

#from django.core.serializers.python import Serializer as PythonSerializer
from django.core.serializers.base import Serializer

class Serializer(Serializer):
    internal_use_only = False
    def start_serialization(self):
        self.workbook = xlwt.Workbook(encoding='utf-8')
        self.sheets={}
        self.dateStyle = xlwt.XFStyle()
        self.dateStyle.num_format_str='YYYY-MM-DD hh:mm'
        self.style = xlwt.XFStyle()
        
    def start_object(self, obj):
        self._current = {}
        self._current_order=[]
        
    def end_object(self, obj):
        model_name = smart_unicode(obj._meta)
        if model_name in self.sheets:
            #we did have a sheet for the model
            sheet = self.sheets[model_name]['sheet']
            row_index = self.sheets[model_name]['row_index']
        else:
            #lets create a sheet for the model
            sheet = self.workbook.add_sheet(model_name)
            row_index=1
            self.sheets[model_name] = {'sheet':sheet, 'row_index':row_index}
            style = xlwt.XFStyle()
            style.bold=True
            sheet.write(0, 0, 'pk', style)
            col_index=1
            #write the title on line 0
            for key in self._current_order:
                sheet.write(0,col_index, key)
                col_index +=1
        
        #write pk in col 0
        sheet.write(row_index, 0, smart_unicode(obj._get_pk_val(), strings_only=True))
        
        #and the rest of the fields after that
        col_index=1
        for key in self._current_order:
            value = self._current[key]
            if isinstance(value, datetime):
                style = self.dateStyle
            else:
                style=self.style
            sheet.write(row_index, col_index, value, style)
            col_index += 1
        row_index += 1
        self.sheets[model_name]['row_index']=row_index
        self._current=None
    
    def handle_field(self, obj, field):
        self._current_order.append(field.name)
        value = field._get_val_from_obj(obj)
        # Protected types (i.e., primitives like None, numbers, dates,
        # and Decimals) are passed through as is. All other values are
        # converted to string first.
        if is_protected_type(value):
            self._current[field.name] = value
        else:
            self._current[field.name] = field.value_to_string(obj)  
    
    def handle_fk_field(self, obj, field):
        self._current_order.append(field.name)
        related = getattr(obj, field.name)
        if related is not None:
            if field.rel.field_name == related._meta.pk.name:
                # Related to remote object via primary key
                related = related._get_pk_val()
            else:
                # Related to remote object via other field
                related = getattr(related, field.rel.field_name)
        self._current[field.name] = smart_unicode(related, strings_only=True)

    def handle_m2m_field(self, obj, field):
        self._current_order.append(field.name)
        if field.creates_table:
            self._current[field.name] = [smart_unicode(related._get_pk_val(), strings_only=True)
                               for related in getattr(obj, field.name).iterator()]
    
    def end_serialization(self):
        #self.workbook.save(self.stream)
        
        filename="dumpdata.xls"
        self.stream.write(u"Result saved in %s" % (filename))
        self.workbook.save(filename) #TODO:hack because of stream making unreadable xls.
        #print u"Result saved in %s" % (filename)

    def getvalue(self):
        if callable(getattr(self.stream, 'getvalue', None)):
            return self.stream.getvalue()

class Deserializer(base.Deserializer):
    
    def __init__(self, stream_or_string, **options):
        models.get_apps()
        super(Deserializer, self).__init__(stream_or_string, **options)
        self.wb = xlrd.open_workbook(stream_or_string.name, formatting_info=True)
        self.row_index=-1
        self.sheet_index=-1
        self.sheet=None
        self.model=None
        self.do_stop_iteration=False
        
        
    def next(self):
        if self.do_stop_iteration:
            raise StopIteration
        if not self.sheet:
            self.sheet_index +=1
            self.sheet = self.wb.sheet_by_index(self.sheet_index)
            self.Model = self._get_model(self.sheet.name)
            self.col_name_key, self.name_col_key = _col_index(self.sheet)
            self.row_index=0
        
        self.row_index +=1
        values = _read_row(self.sheet, self.row_index, self.name_col_key)
        
        
        
        if not 'pk' in values:
            raise base.DeserializationError("<object> row is missing the 'pk' attribute")
        else:
            pk =values['pk']
            try:
                data = {self.Model._meta.pk.attname : self.Model._meta.pk.to_python(pk)}
            except:
                raise base.DeserializationError(
                        "<%s> on line %s have a bad 'pk' attribute of '%s': %r" % (
                        self.sheet.name, self.row_index, values['pk'], values, ))
            
        #TODO: Should check for m2m data and handle it properly
        m2m_data = {}
        
        for key, val in values.items():
            if key!="pk":
                # Get the field from the Model. This will raise a
                # FieldDoesNotExist if, well, the field doesn't exist, which will
                # be propagated correctly.
                field = self.Model._meta.get_field(key)
                if field.rel and isinstance(field.rel, models.ManyToManyRel):
                    raise base.DeserializationError("m2m is found for %s, %s but support is not implemented" % (self.sheet.name, key)) 
                elif field.rel and isinstance(field.rel, models.ManyToOneRel):
                    try:
                        data[field.attname] = self._handle_fk_field(val, field)
                    except:
                        print self.sheet.name
                        raise
                else:
                    data[field.name] = val
        
        next_sheet=False #set to true to move on
        if self.row_index+1 >= self.sheet.nrows:
            #at end of sheet
            next_sheet=True
        else:
            #preview next line
            values = _read_row(self.sheet, self.row_index+1, self.name_col_key)
            if (values[self.name_col_key[0]]=='' and values[self.name_col_key[1]]==''):
                #empty line
                next_sheet=True
            if (values[self.name_col_key[0]]==None):
                raise Exception("Pk is none")
        
        #if we should move to next sheet then do so
        if next_sheet :
            self.sheet=None
            if self.sheet_index+1 >= self.wb.nsheets:
                self.do_stop_iteration=True
        return base.DeserializedObject(self.Model(**data), m2m_data)
    
    def _handle_fk_field(self, value, field):     
         # Check if there is a child node named 'None', returning None if so.
        if value == "none" :
            return None
        else:
            return field.rel.to._meta.get_field(field.rel.field_name).to_python(
                       value)
   
    def _get_model(self, model_identifier):
        """
        Helper to look up a model from an "app_label.module_name" string.
        """
        try:
            Model = models.get_model(*model_identifier.split("."))
        except TypeError:
            Model = None
        if Model is None:
            raise base.DeserializationError(u"Invalid model identifier: '%s'" % model_identifier)
        return Model
             
def _col_index(sheet):
    col_name_key={}
    name_col_key={}
    #first line is titles
    columns = sheet.ncols
    for col_index in range(columns):
        value = str(sheet.cell(0, col_index).value).strip().lower().replace(" ", "_")
        #print "%s = %s" %(value, col_index)
        col_name_key[value]=col_index
        name_col_key[col_index]=value
    return (col_name_key, name_col_key)

def _read_row(sheet, row_index, name_col_key):
    columns = sheet.ncols
    values={}
    for col_index in range(columns):
        #print "%s row=%s, col=%s" % (sheet.name, row_index, col_index,)
        cell = sheet.cell(row_index, col_index)
        if cell.ctype == xlrd.XL_CELL_DATE:
            value = datetime(*xlrd.xldate_as_tuple(cell.value, 0))
        else:
            if isinstance(cell.value, float):
                value="%d" %  (cell.value)
            else:
                value=cell.value
        values[name_col_key[col_index]] = value
    return values
