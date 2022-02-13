using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Structure;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreationModelPlugin
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class CreationModel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            List<Level> listLevel = new FilteredElementCollector(doc)
                        .OfClass(typeof(Level))
                        .OfType<Level>()
                        .ToList();

            Level basedLevel = GetLevel(listLevel, "Уровень 1");
            Level upLevel = GetLevel(listLevel, "Уровень 2");

            List<Wall> walls = CreateWalls(doc, basedLevel, upLevel);
            AddDoor(doc, basedLevel, walls);
            AddWindows(doc, basedLevel, walls);
            AddExtrusionRoof(doc, upLevel, walls);
            return Result.Succeeded;
        }

        private void AddExtrusionRoof(Document doc, Level upLevel, List<Wall> walls)
        {
            LocationCurve curve = walls[1].Location as LocationCurve;
            XYZ startPoint = curve.Curve.GetEndPoint(0);
            XYZ endPoint = curve.Curve.GetEndPoint(1);

            LocationCurve locationCurve = walls[0].Location as LocationCurve;
            var distance = locationCurve.Curve.Length+4;

            RoofType roofType = new FilteredElementCollector(doc)
                        .OfClass(typeof(RoofType))
                        .OfType<RoofType>()
                        .Where(x => x.Name.Equals("Типовой - 125мм"))
                        .FirstOrDefault();

            using (var ts = new Transaction(doc, "Create roof"))
            {
                ts.Start();

                CurveArray curveArray = new CurveArray();
                curveArray.Append(Line.CreateBound(new XYZ(0, startPoint.Y-2, upLevel.Elevation), new XYZ(0, 0, upLevel.Elevation+10)));
                curveArray.Append(Line.CreateBound(new XYZ(0, 0, upLevel.Elevation+10), new XYZ(0, endPoint.Y+2, upLevel.Elevation)));

                ReferencePlane refPlane = doc.Create.NewReferencePlane(new XYZ(0, 0, 0), new XYZ(0, 0, 1), new XYZ(0, 1, 0), doc.ActiveView);
                ExtrusionRoof roof = doc.Create.NewExtrusionRoof(curveArray, refPlane, upLevel, roofType, -distance * 0.5, distance*0.5);

                ts.Commit();
            }
        }

        private void AddWindows(Document doc, Level basedLevel, List<Wall> walls)
        {
            FamilySymbol windowType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0406 x 1220 мм"))
                .Where(x => x.FamilyName.Equals("Фиксированные"))
                .FirstOrDefault();

            using (var ts = new Transaction(doc, "Create windows"))
            {
                ts.Start();
                if (!windowType.IsActive)
                    windowType.Activate();

                for (int i = 1; i < walls.Count; i++)
                {
                    LocationCurve hostcurve = walls[i].Location as LocationCurve;
                    XYZ point1 = hostcurve.Curve.GetEndPoint(0);
                    XYZ point2 = hostcurve.Curve.GetEndPoint(1);
                    XYZ point = (point1 + point2) / 2;
                    FamilyInstance instance = doc.Create.NewFamilyInstance(point, windowType, walls[i], basedLevel, StructuralType.NonStructural);
                    double offset = UnitUtils.ConvertToInternalUnits(900, UnitTypeId.Millimeters);
                    instance.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM).Set(offset);
                }

                ts.Commit();
            }

        }

        private void AddDoor(Document doc, Level basedLevel, List<Wall> walls)
        {
            FamilySymbol doorType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0915 x 2134 мм"))
                .Where(x => x.FamilyName.Equals("Одиночные-Щитовые"))
                .FirstOrDefault();

            LocationCurve hostcurve = walls[0].Location as LocationCurve;
            XYZ point1 = hostcurve.Curve.GetEndPoint(0);
            XYZ point2 = hostcurve.Curve.GetEndPoint(1);
            XYZ point = (point1 + point2) / 2;

            using (var ts = new Transaction(doc, "Create door"))
            {
                ts.Start();

                if (!doorType.IsActive)
                    doorType.Activate();
                doc.Create.NewFamilyInstance(point, doorType, walls[0], basedLevel, StructuralType.NonStructural);

                ts.Commit();
            }
        }

        private List<Wall> CreateWalls(Document doc, Level basedLevel, Level upLevel)
        {
            List<Wall> walls = new List<Wall>();
            List<XYZ> definingPoints = GetDefiningPoints();
            using (Transaction ts = new Transaction(doc, "Create wall"))
            {
                ts.Start();

                for (int i = 0; i < 4; i++)
                {
                    Line line = Line.CreateBound(definingPoints[i], definingPoints[i + 1]);
                    Wall wall = Wall.Create(doc, line, basedLevel.Id, false);
                    walls.Add(wall);
                    wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(upLevel.Id);
                }

                ts.Commit();
            }
            return walls;
        }

        private List<XYZ> GetDefiningPoints()
        {
            double width = UnitUtils.ConvertToInternalUnits(10000, UnitTypeId.Millimeters);
            double depth = UnitUtils.ConvertToInternalUnits(5000, UnitTypeId.Millimeters);
            double dx = width / 2;
            double dy = depth / 2;

            List<XYZ> points = new List<XYZ>();
            points.Add(new XYZ(-dx, -dy, 0));
            points.Add(new XYZ(dx, -dy, 0));
            points.Add(new XYZ(dx, dy, 0));
            points.Add(new XYZ(-dx, dy, 0));
            points.Add(new XYZ(-dx, -dy, 0));

            return points;
        }

        private Level GetLevel(List<Level> listLevel, string nameLevel)
        {
            Level level = listLevel.Where(x => x.Name.Equals(nameLevel)).FirstOrDefault();
            return level;
        }
    }
}
