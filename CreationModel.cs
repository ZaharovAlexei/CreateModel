using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
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

            CreateWalls(doc, basedLevel, upLevel);
            return Result.Succeeded;
        }

        public List<Wall> CreateWalls(Document doc, Level basedLevel, Level upLevel)
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

        public List<XYZ> GetDefiningPoints()
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

        public Level GetLevel(List<Level> listLevel, string nameLevel)
        {
            Level level = listLevel.Where(x => x.Name.Equals(nameLevel)).FirstOrDefault();
            return level;
        }
    }
}
