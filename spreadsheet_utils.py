from oauth2client.service_account import ServiceAccountCredentials
from googleapiclient import discovery
import json
import re

class SpreadsheetObject():
    prefix = "https://sheets.googleapis.com/v4/spreadsheets/"
    def __init__(self,ssid):
        """
        Para obtener el client_secret.json seguir el tutorial de la siguiente web:
        https://github.com/burnash/gspread/blob/master/docs/oauth2.rst
        """
        self.id = ssid
        scope = ['https://spreadsheets.google.com/feeds']
        credentials = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
        self.service = discovery.build('sheets', 'v4', credentials=credentials)
    
    def __str__(self):
        return self.id

    def get_spreadsheet(self):
        """
        Devuelve datos basicos del spreadsheet
        """
        response = self.service.spreadsheets().get(spreadsheetId=self.id).execute()
        return response

    def get_sheets(self):
        """
        Devuelve las paginas
        """
        response = self.get_spreadsheet()
        return response["sheets"]

    def get_sheet(self,index):
        """
        Devuelve una pagina especifica
        """
        response = self.get_spreadsheet()
        try:
            return response["sheets"][index]
        except IndexError:
            return None

    def get_url(self):
        """
        Devuelve la url del spreadsheet
        """
        response = self.get_spreadsheet()
        return response["url"]

    def write_cells(self,range_,values,format_=1, sheetId=None, sheetName=""):
        """
        Escribe en las celdas, el formato de values tiene que ser el de una matriz: [[...],[...],[...]]
        """
        if (sheetId!=None):
            sheet = self.get_sheet(sheetId)
            if not sheet:
                raise Exception("Sheet out of index")
            sheetName = sheet["properties"]["title"]
        if (sheetName):
            range_ = "'" + sheetName + "'!" + range_
        response = self.service.spreadsheets().values().update(spreadsheetId=self.id, range=range_, valueInputOption="RAW", responseValueRenderOption=['FORMATTED_VALUE', 'UNFORMATTED_VALUE', 'FORMULA'][format_], body={"values":values}).execute()
        return response

    def write_cell(self,cell,value,format_=1, sheetId=None, sheetName=""):
        """
        Escribe una unica celda
        """
        if (sheetId!=None):
            sheet = self.get_sheet(sheetId)
            if not sheet:
                raise Exception("Sheet out of index")
            sheetName = sheet["properties"]["title"]
        cell = "'" + sheetName + "'!" + cell
        response = self.service.spreadsheets().values().update(spreadsheetId=self.id, range=cell, valueInputOption="RAW", responseValueRenderOption=['FORMATTED_VALUE', 'UNFORMATTED_VALUE', 'FORMULA'][format_], body={"values":[[value]]}).execute()
        return response

    def clear_cell(self,cell,format_=1, sheetId=None, sheetName=""):
        """
        Borra celdas
        """
        if (sheetId!=None):
            sheet = self.get_sheet(sheetId)
            if not sheet:
                raise Exception("Sheet out of index")
            sheetName = sheet["properties"]["title"]
        cell = "'" + sheetName + "'!" + cell
        response = self.service.spreadsheets().values().clear(spreadsheetId=self.id, range=cell).execute()
        return response

    def get_cell(self,cell,format_=1, sheetId=None, sheetName=""):
        """
        Obtiene celdas
        """
        if (sheetId!=None):
            sheet = self.get_sheet(sheetId)
            if not sheet:
                raise Exception("Sheet out of index")
            sheetName = sheet["properties"]["title"]
        cell = "'" + sheetName + "'!" + cell
        response = self.service.spreadsheets().values().get(spreadsheetId=self.id, range=cell).execute()
        return response["values"]
    
    def format_cells(self, data):
        """
        Recibe en data el objeto FormatObject donde en una unica ejecucion realiza todo lo que tiene el objeto agregado para hacer
        Web donde se puede ver documentacion acerca de este metodo:
        https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/batchUpdate
        Web donde se pueden ver todos los tipos de requests que se pueden agregar:
        https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#Request
        """
        response = self.service.spreadsheets().batchUpdate(spreadsheetId=self.id,body={"requests":data.dict()}).execute()
        return response

class FormatObject():
    def __init__(self):
        self.data = []
    
    def dict(self):
        return self.data
    
        # self.updateCells = None
    
    def _calculate_range(self,first_corner,second_corner,sheet):
        """
        Obtiene dos corners opuestos de una grilla y devuelve el objeto de google GridRange
        """
        start_column = 0
        end_column = 0
        start_row = 0
        end_row = 0
        start_column,start_row = self._get_column_row_values(first_corner)
        end_column,end_row = self._get_column_row_values(second_corner)
        return {
            "startColumnIndex": start_column-1,
            "endColumnIndex": end_column,
            "startRowIndex": start_row-1,
            "endRowIndex": end_row,
            "sheetId": sheet
        }
    
    def _get_column_row_values(self,coordinates):
        """
        Convierte el valor de las celdas en posiciones numericas que necesita el objeto de google GridRange
        """
        coordinates_list = re.split("(\d+)",coordinates)
        column = list(reversed(re.findall("(\w)",coordinates_list[0].lower())))
        row = coordinates_list[1]
        row_value = int(row)
        column_value = 0
        for index,c in enumerate(column):
            column_value = column_value + (ord(c)-96) * (26**index)
        return column_value, row_value

    def set_merged_cells(self,sheet_id,first_corner,second_corner):
        """
        une dentro del rango first_corner y second_corner las celdas
        """
        range_ = self._calculate_range(first_corner,second_corner,sheet_id)
        tmp = {
            "mergeCells":{
                "mergeType": "MERGE_ALL",
                "range": range_
            }
        }
        self.data.append(tmp)
        return tmp

    def set_unmerged_cells(self,sheet_id,first_corner,second_corner):
        """
        desune dentro del rango first_corner y second_corner la celda previamente unida
        """
        range_ = self._calculate_range(first_corner,second_corner,sheet_id)
        tmp = {
            "unmergeCells":{
                "range": range_
            }
        }
        self.data.append(tmp)
        return tmp
    
    def set_basic_filter(self,sheet_id,first_corner,second_corner):
        """
        Agrega en el rango de first_corner y second_corner un filtro
        """
        range_ = self._calculate_range(first_corner,second_corner,sheet_id)
        tmp = {
            "setBasicFilter":{
                "filter":{
                    "range": range_
                }
            }
        }
        self.data.append(tmp)
        return tmp
    
    def set_borders(self,sheet_id,first_corner,second_corner,bottom=None,left=None,right=None,top=None,innerVertical=None,innerHorizontal=None):
        """
        setea los bordes dentro de los limites de first_corner y second_corner, ninguno de los valores del diccionario de abajo son requeridos
        bottom/left/right/top/innerVertical/innerHorizontal = {
            "color": {
                "alpha": 0-1,
                "blue": 0-1,
                "green": 0-1,
                "red": 0-1
            },
            "style": "DOTTED"/"DASHED"/"SOLID"/"SOLID_MEDIUM"/"SOLID_THICK"/"NONE"/"DOUBLE",
            "width": int
        }
        """
        range_ = self._calculate_range(first_corner,second_corner,sheet_id)
        tmp = {
            "updateBorders":{
                "range": range_,
            }
        }
        if bottom:
            tmp["updateBorders"]["bottom"] = bottom
        if left:
            tmp["updateBorders"]["left"] = left
        if right:
            tmp["updateBorders"]["right"] = right
        if top:
            tmp["updateBorders"]["top"] = top
        if innerVertical:
            tmp["updateBorders"]["innerVertical"] = innerVertical
        if innerHorizontal:
            tmp["updateBorders"]["innerHorizontal"] = innerHorizontal
        
        self.data.append(tmp)
        return tmp

    def auto_resize_dimensions(self,sheet_id,start,end,dimension):
        """
        recalcula el tamaño de la dimension seleccionada dentro del rango seleccionado
        dimension = "ROWS"/"COLUMNS"
        """
        tmp = {
            "autoResizeDimensions": {
                "dimensions": {
                    "sheetId": sheet_id,
                    "dimension": dimension,
                    "endIndex": end,
                    "startIndex": start
                }
            }
        }
        self.data.append(tmp)
        return tmp

    def froze_row(self,sheet_id,row_index):
        """
        Este metodo congela las primeras row_index filas del excel
        """
        tmp = {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {
                        "frozenRowCount": row_index
                    }
                },
                "fields": "gridProperties.frozenRowCount"
            }
        }
        self.data.append(tmp)
        return tmp

    def set_cell_format(self,sheet_id,first_corner,second_corner,format_):
        """
        Dentro de los limites de first_corner y second_corner, da el formato de format_
        ninguno de los valores de format_ es requerido y no es necesario armar el diccionario completo
        format_ = {
            "backgroundColor": {
                "alpha": 0-1,
                "blue": 0-1,
                "green": 0-1,
                "red": 0-1
            },
            "borders": {
                "top": {
                    "color": {
                        "alpha": 0-1,
                        "blue": 0-1,
                        "green": 0-1,
                        "red": 0-1
                    },
                    "style": "DOTTED"/"DASHED"/"SOLID"/"SOLID_MEDIUM"/"SOLID_THICK"/"NONE"/"DOUBLE",
                    "width": int
                },
                "right": {
                    "color": {
                        "alpha": 0-1,
                        "blue": 0-1,
                        "green": 0-1,
                        "red": 0-1
                    },
                    "style": "DOTTED"/"DASHED"/"SOLID"/"SOLID_MEDIUM"/"SOLID_THICK"/"NONE"/"DOUBLE",
                    "width": int                    
                },
                "left": {
                    "color": {
                        "alpha": 0-1,
                        "blue": 0-1,
                        "green": 0-1,
                        "red": 0-1
                    },
                    "style": "DOTTED"/"DASHED"/"SOLID"/"SOLID_MEDIUM"/"SOLID_THICK"/"NONE"/"DOUBLE",
                    "width": int
                },
                "bottom": {
                    "color": {
                        "alpha": 0-1,
                        "blue": 0-1,
                        "green": 0-1,
                        "red": 0-1
                    },
                    "style": "DOTTED"/"DASHED"/"SOLID"/"SOLID_MEDIUM"/"SOLID_THICK"/"NONE"/"DOUBLE",
                    "width": int
                }
            },
            "horizontalAlignment": "LEFT"/"CENTER"/"RIGHT",
            "hyperlinkDisplayType": "LINKED"/"PLAIN_TEXT",
            "numberFormat": {
                "type": "TEXT"/"NUMBER"/"CURRENCY"/"DATE"/"SCIENTIFIC"/"TIME"/"DATE_TIME",
                "pattern": ""
            },
            "padding": {
                "bottom": int,
                "left": int,
                "right": int,
                "top": int
            },
            "textDirection": "LEFT_TO_RIGHT"/"RIGHT_TO_LEFT",
            "textFormat": {
                "bold": bool,
                "fontFamily": "",
                "fontSize": int,
                "foregroundColor": {
                    "alpha": 0-1,
                    "blue": 0-1,
                    "green": 0-1,
                    "red": 0-1
                },
                "italic": bool,
                "strikethrough": bool,
                "underline": bool
            },
            "textRotation": {
                "angle": int,
                "vertical": bool
            },
            "verticalAlignment": "TOP"/"MIDDLE"/"BOTTOM",
            "wrapStrategy": "WRAP"/"CLIP"/"LEGACY_WRAP"/"OVERFLOW_CELL"
        }
        """
        range_ = self._calculate_range(first_corner,second_corner,sheet_id)
        tmp = {
            "repeatCell": {
                "fields": "*",
                "cell": {
                    "userEnteredFormat": format_
                },
                "range": range_
            }
        }
        self.data.append(tmp)
        return tmp
    
    def set_raw(self,raw):
        """
        Usar bajo propia precaucion, aca se puede agregar un request extra customizado
        """
        self.data.append(raw)
        return raw


sso = SpreadsheetObject("1To7sf10-ehkbWQ-3mHOiZG3vfVbe__t3viP8_rKx-tA")

fo = FormatObject()

fo.set_raw({
      "unmergeCells": {
        "range": {
          "endColumnIndex": 10,
          "sheetId": 0,
          "endRowIndex": 5,
          "startColumnIndex": 0,
          "startRowIndex": 0
        }
      }})

sso.format_cells(fo)


