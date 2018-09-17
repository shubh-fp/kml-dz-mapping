from fastkml import kml
from shapely.geometry import Point, Polygon
import xlrd
import xlsxwriter

latColumn = None
lngColumn = None
# listOfVendorPoints = []
listOfPolygons = []
vendorCodeAndPointsMap = {}
foundPolygonSet = set()

vendorCodeHeader = 'VendorCode'
latitudeHeader = 'Latitude'
longitudeHeader = 'Longitude'
isDeliveryZoneHeader = 'Falling in a deliery zone'
dzNameHeader = 'Dzname'
distanceBoundaryHeader = 'Distance from nearest boundary'
nearestDeliveryZoneHeader = 'Nearest delivery zone'
distanceFromDzHeader = 'Distance from nearest dz'

def findNearestPolygon(f2, p):
	minDistance = float("inf")
	vendorName = None
	for placemark in f2:
		if placemark._geometry.geometry.geom_type == 'Polygon':
			distance = placemark._geometry.geometry.exterior.distance(p)
			if distance < minDistance:
				minDistance = distance
				vendorName = placemark.name
	return (vendorName, minDistance)

wb = xlrd.open_workbook('latlng.xlsx')
sheet = wb.sheet_by_index(0)

workBook = xlsxwriter.Workbook('demo.xlsx')
workSheet = workBook.add_worksheet()
bold = workBook.add_format({'bold': True})

# Write Headers
workSheet.set_column('A:A', len(vendorCodeHeader))
workSheet.write('A1', vendorCodeHeader, bold)
workSheet.set_column('B:B', len(latitudeHeader)+5)
workSheet.write('B1', latitudeHeader, bold)
workSheet.set_column('C:C', len(longitudeHeader)+5)
workSheet.write('C1', longitudeHeader, bold)
workSheet.set_column('D:D', len(isDeliveryZoneHeader))
workSheet.write('D1', isDeliveryZoneHeader, bold)
workSheet.set_column('E:E', len(dzNameHeader)+40)
workSheet.write('E1', dzNameHeader, bold)
workSheet.set_column('F:F', len(distanceBoundaryHeader))
workSheet.write('F1', distanceBoundaryHeader, bold)
workSheet.set_column('G:G', len(nearestDeliveryZoneHeader)+20)
workSheet.write('G1', nearestDeliveryZoneHeader, bold)
workSheet.set_column('H:H', len(distanceFromDzHeader))
workSheet.write('H1', distanceFromDzHeader, bold)

# Find lat and lng column
for i in range(sheet.ncols):
	columnHeader = sheet.cell_value(0,i).lower()
	if columnHeader == 'lat' or columnHeader == 'latitude':
		latColumn = i
	if columnHeader == 'lng' or columnHeader == 'lon' or columnHeader == 'longitude':
		lngColumn = i

print ('Column found for latitude: ' , latColumn)
print ('Column found for longitude: ' , lngColumn)

# Make Point objects from sheet
for i in range(sheet.nrows-1):
	vendorCodeAndPointsMap[sheet.cell_value(i+1, 0)] = Point(sheet.cell_value(i+1, lngColumn), sheet.cell_value(i+1, latColumn))
	# listOfVendorPoints.append(Point(sheet.cell_value(i+1, lngColumn), sheet.cell_value(i+1, latColumn)))

with open('Zones.kml') as kmlFile:
	doc = kmlFile.read()

k = kml.KML()
k.from_string(doc)
f = list(k.features())
f2 = list(f[0].features())
rowA = 1
columnA = 0

for key, value in vendorCodeAndPointsMap.items():
	for placemark in f2:
		if placemark._geometry.geometry.geom_type == 'Polygon':
			if (placemark._geometry.geometry.contains(value)):
				foundPolygonSet.add(key)
				nearestEdgeDist = placemark._geometry.geometry.exterior.distance(value)
				rowData = [key, value.y, value.x, 'Y', placemark.name, nearestEdgeDist, 'NA', 'NA']
				for val in rowData:
					workSheet.write(rowA, columnA, val)
					columnA += 1
				rowA += 1
				columnA = 0


for key, value in vendorCodeAndPointsMap.items():
	if(key not in foundPolygonSet):
		(vendorName, nearestPolyDist) = findNearestPolygon(f2, value)
		rowData = [key, value.y, value.x, 'N', 'NA', 'NA', vendorName, nearestPolyDist]
		for val in rowData:
			workSheet.write(rowA, columnA, val)
			columnA += 1
		rowA += 1
		columnA = 0

workBook.close()