using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace CreationModel
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class CreationModel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uIDocument = commandData.Application.ActiveUIDocument;
            Document document = uIDocument.Document;

            Level level1 = GetLevelByName(document, "Уровень 1");
            Level level2 = GetLevelByName(document, "Уровень 2");

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

            List<Wall> walls = new List<Wall>();

            Transaction transaction = new Transaction(document);
            transaction.Start("Создание стен");
            for (int i = 0; i < 4; i++)
            {
                Wall wall = CreateWallByPoints(document, points[i], points[i + 1], level1.Id, level2.Id);
                walls.Add(wall);
                if (i == 0)
                    AddDoorOrWindow(document, level1, walls[0], BuiltInCategory.OST_Doors, "0915 x 2134 мм", "Одиночные-Щитовые");
                else
                    AddDoorOrWindow(document, level1, walls[i], BuiltInCategory.OST_Windows, "0915 x 1830 мм", "Фиксированные");

            }
            //AddRoof(document, level2, walls);
            AddExtrusionRoof(document, level2, walls,  dx,  dy);

            transaction.Commit();

            return Result.Succeeded;
        }

        public void AddExtrusionRoof(Document document, Level level2, List<Wall> walls, double dx, double dy)
        {
            RoofType roofType = new FilteredElementCollector(document)
                     .OfClass(typeof(RoofType))
                     .OfType<RoofType>()
                     .Where(x => x.Name.Equals("Типовой - 400мм"))
                     .Where(x => x.FamilyName.Equals("Базовая крыша"))
            .FirstOrDefault();

            CurveArray curveArray = new CurveArray();
            LocationCurve curve = walls[1].Location as LocationCurve;
            LocationCurve endCurve = walls[0].Location as LocationCurve;
            double extrEndCurve = endCurve.Curve.Length;
            double dh = walls[0].get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble();

            List<XYZ> points = new List<XYZ>();
            points.Add(new XYZ(-dx, -dy, dh));
            points.Add(new XYZ(dx, -dy, dh));
            points.Add(new XYZ(dx, dy, dh));
            points.Add(new XYZ(-dx, dy, dh));
            points.Add(new XYZ(-dx, -dy, dh));

            curveArray.Append(Line.CreateBound(points[1], (points[1] + points[2])/2 + new XYZ(0, 0, dh/2)));
            curveArray.Append(Line.CreateBound((points[1] + points[2]) / 2 + new XYZ(0, 0, dh / 2), points[2]));

            ReferencePlane plane = document.Create.NewReferencePlane(new XYZ(0, 0, 0), new XYZ(0, 0, dh / 2), new XYZ(0, dy, 0), document.ActiveView);
            document.Create.NewExtrusionRoof(curveArray, plane,level2,roofType, dx, -dx);
        }

        public void AddRoof(Document document, Level level2, List<Wall> walls)
        {
            RoofType roofType = new FilteredElementCollector(document)
                     .OfClass(typeof(RoofType))
                     .OfType<RoofType>()
                     .Where(x => x.Name.Equals("Типовой - 400мм"))
                     .Where(x => x.FamilyName.Equals("Базовая крыша"))
                     .FirstOrDefault();
            double wallWidth = walls[0].Width;
            double dt = wallWidth / 2;
            List<XYZ> points = new List<XYZ>();
            points.Add(new XYZ(-dt, -dt, 0));
            points.Add(new XYZ(dt, -dt, 0));
            points.Add(new XYZ(dt, dt, 0));
            points.Add(new XYZ(-dt, dt, 0));
            points.Add(new XYZ(-dt, -dt, 0));

            Application application = document.Application;

            CurveArray curveArray = application.Create.NewCurveArray();
            for (int i = 0; i < walls.Count; i++)
            {
                LocationCurve curve = walls[i].Location as LocationCurve;
                XYZ point1 = curve.Curve.GetEndPoint(0);
                XYZ point2 = curve.Curve.GetEndPoint(1);
                Line line = Line.CreateBound(point1 + points[i], point2 + points[i+1]);
                curveArray.Append(line);
            }
             
            ModelCurveArray modelCurveArray = new ModelCurveArray();
            FootPrintRoof footPrintRoof = document.Create.NewFootPrintRoof(curveArray, level2, roofType, out modelCurveArray);

            ModelCurveArrayIterator iterator = modelCurveArray.ForwardIterator();
            iterator.Reset();

            while (iterator.MoveNext())
            {
                ModelCurve modelCurve = iterator.Current as ModelCurve;
                footPrintRoof.set_DefinesSlope(modelCurve, true);
                footPrintRoof.set_SlopeAngle(modelCurve, 0.5);
            }
        }

        public void AddDoorOrWindow(Document document, Level level1, Wall wall, BuiltInCategory builtInCategory, string typeName, string familyName)
        {
            FamilySymbol familySymbol = new FilteredElementCollector(document)
                      .OfClass(typeof(FamilySymbol))
                      .OfCategory(builtInCategory)
                      .OfType<FamilySymbol>()
                      .Where(x => x.Name.Equals(typeName))
                      .Where(x => x.FamilyName.Equals(familyName))
                      .FirstOrDefault();

            if (!familySymbol.IsActive)
                familySymbol.Activate();
    
            LocationCurve curve = wall.Location as LocationCurve;
            XYZ point1 = curve.Curve.GetEndPoint(0);
            XYZ point2 = curve.Curve.GetEndPoint(1);
            XYZ point = (point1 + point2) / 2;

            document.Create.NewFamilyInstance(point, familySymbol, wall, level1, StructuralType.NonStructural);
        }

        
        public Level GetLevelByName(Document document, string name)
        {
            Level level = new FilteredElementCollector(document)
                     .OfClass(typeof(Level))
                     .OfType<Level>()
                     .Where(x => x.Name.Equals(name))
                     .FirstOrDefault();
            return level;

        }
       
        public Wall CreateWallByPoints(Document document, XYZ point1, XYZ point2, ElementId level1Id, ElementId level2Id)
        {
            Line line = Line.CreateBound(point1, point2);
            Wall wall = Wall.Create(document, line, level1Id, false);
            wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(level2Id);
            return wall;
        }
    }



}
