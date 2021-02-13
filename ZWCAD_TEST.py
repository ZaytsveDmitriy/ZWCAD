# -*- coding: utf8 -*-
from ZWCAD_CONSTANTS import constants

from win32com import client
import pythoncom
import array
import pythoncom
zcad = client.Dispatch("ZWCAD.Application")
doc = zcad.ActiveDocument
model = doc.ModelSpace
excel = client.Dispatch("Excel.Application")
wb = excel.Workbooks.Open(u'C:\DATA\PROGRAMS\ZWCAD\ИД.xlsx')
sheet = wb.ActiveSheet

def make_point(x, y, z = 0.0):
	# функция нужна только для питона
	return client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8,[x, y, z])


print (model.Count)

objects_istances = model

selection_set = doc.ActiveSelectionSet
print(f"Selection name is {selection_set.Name} and ther is {selection_set.Count} item in" )
for i in range(selection_set.Count):
	items = selection_set.Item(i)
	print(items.TextString, items.ObjectID)

# for object_instance in objects_istances:
# 	print(object_instance.ObjectName)
# 	print(object_instance.ObjectID)

	# for attribute in object_instance.GetAttributes():
	# 	print(attribute.TagString, " ", attribute.TextString, " ", attribute.FieldLength, " ", attribute.Thickness, " ", attribute.Height)
	# 	text = attribute
	# 	print(text)

# i = 0 
# for cel in [cel[0].value for cel in sheet.Range("A1:A10")]: 
# 	pnt_start = make_point(50, 100 + i) # создаем координаты точек
# 	text = model.AddMText(pnt_start , 20, cel) 
# 	text.Height = 30
# 	text.AttachmentPoint = constants.zcAttachmentPointMiddleRight
# 	i +=50

# i = 0 
# for cel in [cel[0].value for cel in sheet.Range("A1:A10")]:
# 	pnt_start = make_point(0, 0 + i) # create point coordinate
# 	block = model.InsertBlock(pnt_start, "SIMPLE_BLOCK", 1.0, 1.0, 1.0, 0) # insert block and get it instance object 
# 	print(block.ObjectID)
# 	for attribute in block.GetAttributes():
# 		# test attribute handling
# 		print(attribute.TagString, " ", attribute.TextString, " ", attribute.FieldLength, " ", attribute.Thickness )
# 		attribute.Thickness = 10
# 		attribute.FieldLength = 10
# 		attribute.TextString = "blablabla yiuyoiuyuio ihiouhkljhkljh"
# 		print(attribute.TagString, " ", attribute.TextString, " ", attribute.FieldLength, " ", attribute.Thickness )
# 		attribute.Update()
# 	i +=300




