using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;

namespace ColumnDimensioner
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    internal class ColumnDimensionerCommand : IExternalCommand
    {
        private const double Eps = 1e-6;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                try
                {
                    _ = GetPluginStartInfo();
                }
                catch
                {
                }

                UIDocument uiDoc = commandData.Application.ActiveUIDocument;
                Document doc = uiDoc.Document;
                Selection sel = uiDoc.Selection;
                View activeView = doc.ActiveView;

                if (activeView.ViewType != ViewType.FloorPlan && activeView.ViewType != ViewType.EngineeringPlan)
                {
                    TaskDialog.Show("Revit", "Перед запуском плагина откройте \"План этажа\" или \"План несущих конструкций\"");
                    return Result.Cancelled;
                }

                List<DimensionType> dimensionTypesList = new FilteredElementCollector(doc)
                    .OfClass(typeof(DimensionType))
                    .Cast<DimensionType>()
                    .Where(dt => dt.StyleType == DimensionStyleType.Linear)
                    .Where(dt => dt.Name != "Стиль линейных размеров")
                    .OrderBy(dt => dt.Name, new AlphanumComparatorFastString())
                    .ToList();
                if (dimensionTypesList.Count == 0)
                {
                    TaskDialog.Show("Revit", "Не удалось найти ни одного типа для линейных размеров!");
                    return Result.Cancelled;
                }

                ColumnDimensionerWPF window = new ColumnDimensionerWPF(dimensionTypesList);
                window.ShowDialog();
                if (window.DialogResult != true)
                {
                    return Result.Cancelled;
                }

                List<FamilyInstance> columnsList = GetColumns(doc, sel, activeView, window);
                if (columnsList.Count == 0)
                {
                    TaskDialog.Show("Revit", "Не выбрано ни одной колонны!");
                    return Result.Cancelled;
                }

                DimensionType dimensionType = window.SelectedDimensionType;
                if (dimensionType == null)
                {
                    TaskDialog.Show("Revit", "Не удалось определить тип размера.");
                    return Result.Cancelled;
                }

                bool firstRowEnabled = window.IndentationFirstRowDimensionsIsChecked;
                double firstRowOffset = 0.0;
                if (firstRowEnabled &&
                    !TryParseOffsetInMillimeters(window.IndentationFirstRowDimensions, out firstRowOffset))
                {
                    TaskDialog.Show("Revit", "Некорректный отступ первого ряда размеров. Введите число в мм.");
                    return Result.Cancelled;
                }

                bool secondRowEnabled = window.IndentationSecondRowDimensionsIsChecked;
                double secondRowOffset = 0.0;
                if (secondRowEnabled &&
                    !TryParseOffsetInMillimeters(window.IndentationSecondRowDimensions, out secondRowOffset))
                {
                    TaskDialog.Show("Revit", "Некорректный отступ второго ряда размеров. Введите число в мм.");
                    return Result.Cancelled;
                }

                int scale = GetViewScale(activeView);

                List<Grid> gridsList = new FilteredElementCollector(doc, activeView.Id)
                    .OfCategory(BuiltInCategory.OST_Grids)
                    .WhereElementIsNotElementType()
                    .Cast<Grid>()
                    .ToList();

                Options opt = new Options
                {
                    ComputeReferences = true,
                    IncludeNonVisibleObjects = false,
                    View = activeView
                };

                using (TransactionGroup tg = new TransactionGroup(doc, "Образмерить колонны"))
                {
                    tg.Start();

                    TextNoteType tempTextType;
                    bool createdTempTextType;
                    if (!TryGetOrCreateTempTextType(doc, dimensionType, out tempTextType, out createdTempTextType))
                    {
                        tg.RollBack();
                        TaskDialog.Show("Revit", "Не удалось создать временный тип текста.");
                        return Result.Failed;
                    }

                    List<string> errors = new List<string>();

                    foreach (FamilyInstance column in columnsList)
                    {
                        try
                        {
                            ProcessColumn(
                                doc,
                                activeView,
                                column,
                                dimensionType,
                                tempTextType,
                                gridsList,
                                opt,
                                scale,
                                firstRowEnabled,
                                firstRowOffset,
                                secondRowEnabled,
                                secondRowOffset,
                                errors);
                        }
                        catch (Exception ex)
                        {
                            string colId = IdToText(column?.Id);
                            errors.Add($"Колонна {colId}: {ex.Message}");
                        }
                    }

                    using (Transaction t = new Transaction(doc, "Удалить временный тип текста"))
                    {
                        t.Start();
                        if (createdTempTextType && tempTextType != null && doc.GetElement(tempTextType.Id) != null)
                        {
                            ElementId id = tempTextType.Id;
                            doc.Delete(id);
                        }
                        t.Commit();
                    }

                    tg.Assimilate();

                    if (errors.Count > 0)
                    {
                        string text = string.Join(Environment.NewLine, errors.Take(20));
                        if (errors.Count > 20)
                        {
                            text += Environment.NewLine + "...";
                        }

                        TaskDialog.Show("ColumnDimensioner",
                            "Часть колонн не удалось обработать:" + Environment.NewLine + text);
                    }
                }

                TaskDialog.Show("ColumnDimensioner", "Обработка завершена.");
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("ColumnDimensioner", "Ошибка выполнения внешней команды.");
                return Result.Failed;
            }
        }
        private static void ProcessColumn(
            Document doc,
            View activeView,
            FamilyInstance column,
            DimensionType dimensionType,
            TextNoteType tempTextType,
            List<Grid> gridsList,
            Options opt,
            int scale,
            bool firstRowEnabled,
            double firstRowOffset,
            bool secondRowEnabled,
            double secondRowOffset,
            List<string> errors)
        {
            if (column == null || !TryGetColumnLocationPoint(column, out XYZ columnLocationPoint))
            {
                return;
            }

            string columnId = IdToText(column.Id);

            XYZ hand = NormalizeOrNull(column.HandOrientation);
            XYZ facing = NormalizeOrNull(column.FacingOrientation);

            if (hand == null || facing == null)
            {
                errors.Add($"Колонна {columnId}: не удалось получить HandOrientation/FacingOrientation.");
                return;
            }

            if (!TryGetColumnDimensions(column, opt, hand, facing, out double width, out double height))
            {
                errors.Add($"Колонна {columnId}: не удалось определить габариты.");
                return;
            }

            ProcessColumnAxis(
                doc,
                activeView,
                column,
                dimensionType,
                tempTextType,
                gridsList,
                opt,
                scale,
                firstRowEnabled,
                firstRowOffset,
                secondRowEnabled,
                secondRowOffset,
                errors,
                columnId,
                columnLocationPoint,
                measureAxis: hand,
                offsetAxis: facing.Negate(),
                gridDirection: facing,
                dimensionLength: width,
                offsetDistanceForSecondRow: height,
                offsetDistanceForFirstRow: height,
                dirIndex: 0,
                mainTransactionName: "Размер вдоль HandOrientation",
                secondaryTransactionName: "Размер вдоль HandOrientation второстепенный",
                noGridMessage: "для HandOrientation не найдена подходящая ось.");

            ProcessColumnAxis(
                doc,
                activeView,
                column,
                dimensionType,
                tempTextType,
                gridsList,
                opt,
                scale,
                firstRowEnabled,
                firstRowOffset,
                secondRowEnabled,
                secondRowOffset,
                errors,
                columnId,
                columnLocationPoint,
                measureAxis: facing,
                offsetAxis: hand,
                gridDirection: hand,
                dimensionLength: height,
                offsetDistanceForSecondRow: width,
                offsetDistanceForFirstRow: width,
                dirIndex: 1,
                mainTransactionName: "Размер вдоль FacingOrientation",
                secondaryTransactionName: "Размер вдоль FacingOrientation второстепенный",
                noGridMessage: "для FacingOrientation не найдена подходящая ось.");
        }


        private static void ProcessColumnAxis(
            Document doc,
            View activeView,
            FamilyInstance column,
            DimensionType dimensionType,
            TextNoteType tempTextType,
            List<Grid> gridsList,
            Options opt,
            int scale,
            bool firstRowEnabled,
            double firstRowOffset,
            bool secondRowEnabled,
            double secondRowOffset,
            List<string> errors,
            string columnId,
            XYZ columnLocationPoint,
            XYZ measureAxis,
            XYZ offsetAxis,
            XYZ gridDirection,
            double dimensionLength,
            double offsetDistanceForSecondRow,
            double offsetDistanceForFirstRow,
            int dirIndex,
            string mainTransactionName,
            string secondaryTransactionName,
            string noGridMessage)
        {
            Line dimensionLine = Line.CreateBound(
                columnLocationPoint + dimensionLength / 2.0 * measureAxis.Negate(),
                columnLocationPoint + dimensionLength / 2.0 * measureAxis);

            ReferenceArray columnRefs = new ReferenceArray();
            if (TryGetColumnReferences(
                column,
                opt,
                dirIndex,
                NormalizeOrNull(column.HandOrientation),
                NormalizeOrNull(column.FacingOrientation),
                out Reference ref1,
                out Reference ref2))
            {
                columnRefs.Append(ref1);
                columnRefs.Append(ref2);
            }

            if (columnRefs.Size < 2)
            {
                errors.Add($"Колонна {columnId}: не удалось получить 2 грани для dir={dirIndex}.");
                return;
            }

            if (secondRowEnabled)
            {
                Dimension mainDimension = null;
                using (Transaction t = new Transaction(doc, mainTransactionName))
                {
                    t.Start();
                    try
                    {
                        mainDimension = doc.Create.NewDimension(
                            activeView,
                            dimensionLine,
                            columnRefs,
                            dimensionType);

                        XYZ translation =
                            offsetDistanceForSecondRow * offsetAxis
                            + secondRowOffset * offsetAxis;

                        ElementTransformUtils.MoveElement(doc, mainDimension.Id, translation);
                        t.Commit();
                    }
                    catch (Exception ex)
                    {
                        errors.Add($"Колонна {columnId}: ошибка создания основного размера dir={dirIndex}. {ex.Message}");
                        t.RollBack();
                    }
                }

                if (mainDimension == null)
                {
                    errors.Add($"Колонна {columnId}: не удалось создать основной размер dir={dirIndex}.");
                }
                else
                {
                    AdjustSingleDimensionText(
                        doc,
                        activeView.Id,
                        tempTextType.Id,
                        mainDimension,
                        dimensionLine.Length,
                        measureAxis,
                        scale);
                }
            }

            if (!firstRowEnabled)
            {
                return;
            }

            if (!TryFindBestGridReference(
                gridsList,
                columnLocationPoint,
                gridDirection,
                opt,
                out Reference gridReference,
                out XYZ projectedPointOnGrid,
                out double signedDistance))
            {
                return;
            }

            double pFace1 = GetReferenceProjection(doc, columnRefs.get_Item(0), measureAxis);
            double pFace2 = GetReferenceProjection(doc, columnRefs.get_Item(1), measureAxis);
            double pGrid = projectedPointOnGrid != null
                ? projectedPointOnGrid.DotProduct(measureAxis)
                : GetReferenceProjection(doc, gridReference, measureAxis);

            double minP = Math.Min(pGrid, Math.Min(pFace1, pFace2));
            double maxP = Math.Max(pGrid, Math.Max(pFace1, pFace2));
            double margin = 100.0 / 304.8;

            XYZ originShift = columnLocationPoint - measureAxis * columnLocationPoint.DotProduct(measureAxis);

            Line lineToGrid = Line.CreateBound(
                measureAxis * (minP - margin) + originShift,
                measureAxis * (maxP + margin) + originShift);

            ReferenceArray refs = CreateSortedReferenceArrayByProjection(
                new[]
                {
                    (columnRefs.get_Item(0), pFace1),
                    (columnRefs.get_Item(1), pFace2),
                    (gridReference, pGrid)
                });

            Dimension secondDimension = null;
            using (Transaction t = new Transaction(doc, secondaryTransactionName))
            {
                t.Start();
                try
                {
                    secondDimension = doc.Create.NewDimension(
                        activeView,
                        lineToGrid,
                        refs,
                        dimensionType);

                    XYZ translation =
                        offsetDistanceForFirstRow * offsetAxis
                        + firstRowOffset * offsetAxis;

                    ElementTransformUtils.MoveElement(doc, secondDimension.Id, translation);
                    TryLockDimension(secondDimension);
                    t.Commit();
                }
                catch (Exception ex)
                {
                    errors.Add($"Колонна {columnId}: ошибка создания второстепенного размера dir={dirIndex}. {ex.Message}");
                    t.RollBack();
                }
            }

            if (secondDimension == null)
            {
                errors.Add($"Колонна {columnId}: не удалось создать второстепенный размер dir={dirIndex}.");
            }
            else
            {
                MoveSegmentTexts(
                    doc,
                    activeView.Id,
                    tempTextType.Id,
                    secondDimension,
                    measureAxis,
                    scale);
            }
        }

        private static bool TryFindBestGridReference(
            List<Grid> gridsList,
            XYZ originPoint,
            XYZ expectedGridDirection,
            Options opt,
            out Reference gridReference,
            out XYZ projectedPointOnGrid,
            out double signedDistance)
        {
            gridReference = null;
            projectedPointOnGrid = null;
            signedDistance = 0.0;

            XYZ expectedDir = NormalizeOrNull(expectedGridDirection);
            if (originPoint == null || expectedDir == null)
            {
                return false;
            }

            double bestAbsDistance = double.MaxValue;

            foreach (Grid grid in gridsList)
            {
                if (grid == null || grid.Curve == null)
                {
                    continue;
                }

                if (!TryGetGridCurveDirection(grid.Curve, out XYZ gridDir))
                {
                    continue;
                }

                double dirDot = Math.Abs(gridDir.DotProduct(expectedDir));
                if (dirDot < 0.975)
                {
                    continue;
                }

                IntersectionResult projection = null;
                try
                {
                    projection = grid.Curve.Project(originPoint);
                }
                catch
                {
                }

                if (projection == null || projection.XYZPoint == null)
                {
                    continue;
                }

                XYZ projected = projection.XYZPoint;
                XYZ delta = projected - originPoint;

                XYZ normal = new XYZ(-gridDir.Y, gridDir.X, 0.0);
                normal = NormalizeOrNull(normal);
                if (normal == null)
                {
                    continue;
                }

                double dist = delta.DotProduct(normal);
                double absDist = Math.Abs(dist);

                Reference candidateReference = TryGetGridReference(grid, opt);
                if (candidateReference == null)
                {
                    continue;
                }

                if (absDist < bestAbsDistance)
                {
                    bestAbsDistance = absDist;
                    gridReference = candidateReference;
                    projectedPointOnGrid = projected;
                    signedDistance = dist;
                }
            }

            return gridReference != null && projectedPointOnGrid != null;
        }

        private static Reference TryGetGridReference(Grid grid, Options opt)
        {
            if (grid == null)
            {
                return null;
            }

            try
            {
                // Для устойчивой привязки к оси нужен datum-reference самой оси.
                return new Reference(grid);
            }
            catch
            {
            }

            try
            {
                GeometryElement geom = grid.get_Geometry(opt);
                Reference geometryReference = TryFindReferenceInGeometry(geom);
                if (geometryReference != null)
                {
                    return geometryReference;
                }
            }
            catch
            {
            }

            try
            {
                if (grid.Curve != null && grid.Curve.Reference != null)
                {
                    return grid.Curve.Reference;
                }
            }
            catch
            {
            }

            return null;
        }

        private static Reference TryFindReferenceInGeometry(GeometryElement geometry)
        {
            if (geometry == null)
            {
                return null;
            }

            foreach (GeometryObject obj in geometry)
            {
                if (obj == null)
                {
                    continue;
                }

                if (obj is Curve curve && curve.Reference != null)
                {
                    return curve.Reference;
                }

                if (obj is GeometryInstance gi)
                {
                    try
                    {
                        Reference nestedRef = TryFindReferenceInGeometry(gi.GetInstanceGeometry());
                        if (nestedRef != null)
                        {
                            return nestedRef;
                        }
                    }
                    catch
                    {
                    }

                    try
                    {
                        Reference nestedRef = TryFindReferenceInGeometry(gi.GetSymbolGeometry());
                        if (nestedRef != null)
                        {
                            return nestedRef;
                        }
                    }
                    catch
                    {
                    }
                }
            }

            return null;
        }

        private static bool TryGetGridCurveDirection(Curve curve, out XYZ direction)
        {
            direction = null;
            if (curve == null)
            {
                return false;
            }

            if (curve is Line line)
            {
                direction = NormalizeOrNull(line.Direction);
                return direction != null;
            }

            try
            {
                XYZ p0 = curve.GetEndPoint(0);
                XYZ p1 = curve.GetEndPoint(1);
                direction = NormalizeOrNull(p1 - p0);
                return direction != null;
            }
            catch
            {
                return false;
            }
        }

        private static bool TryGetColumnLocationPoint(FamilyInstance column, out XYZ locationPoint)
        {
            locationPoint = null;

            if (column == null || column.Location == null)
            {
                return false;
            }

            if (column.Location is LocationPoint lp && lp.Point != null)
            {
                locationPoint = lp.Point;
                return true;
            }

            if (column.Location is LocationCurve lc && lc.Curve != null)
            {
                locationPoint = lc.Curve.Evaluate(0.5, true);
                return locationPoint != null;
            }

            return false;
        }

        private static bool TryGetColumnReferences(
            FamilyInstance column,
            Options opt,
            int dir,
            XYZ handOrientation,
            XYZ facingOrientation,
            out Reference ref1,
            out Reference ref2)
        {
            ref1 = null;
            ref2 = null;

            XYZ axis = dir == 0 ? handOrientation : facingOrientation;
            axis = NormalizeOrNull(axis);
            if (axis == null)
            {
                return false;
            }

            try
            {
                if (dir == 0)
                {
                    IList<Reference> leftRefs = column.GetReferences(FamilyInstanceReferenceType.Left);
                    IList<Reference> rightRefs = column.GetReferences(FamilyInstanceReferenceType.Right);

                    if (leftRefs != null && leftRefs.Count > 0 && rightRefs != null && rightRefs.Count > 0)
                    {
                        ref1 = leftRefs.FirstOrDefault();
                        ref2 = rightRefs.FirstOrDefault();
                        if (ref1 != null && ref2 != null)
                        {
                            return true;
                        }
                    }
                }
                else
                {
                    IList<Reference> frontRefs = column.GetReferences(FamilyInstanceReferenceType.Front);
                    IList<Reference> backRefs = column.GetReferences(FamilyInstanceReferenceType.Back);

                    if (frontRefs != null && frontRefs.Count > 0 && backRefs != null && backRefs.Count > 0)
                    {
                        ref1 = frontRefs.FirstOrDefault();
                        ref2 = backRefs.FirstOrDefault();
                        if (ref1 != null && ref2 != null)
                        {
                            return true;
                        }
                    }
                }
            }
            catch
            {
            }

            try
            {
                IList<Reference> strongRefs = column.GetReferences(FamilyInstanceReferenceType.StrongReference);
                if (TryGetOppositeReferencesFromRefSet(column, strongRefs, axis, out ref1, out ref2))
                {
                    return true;
                }
            }
            catch
            {
            }

            try
            {
                IList<Reference> weakRefs = column.GetReferences(FamilyInstanceReferenceType.WeakReference);
                if (TryGetOppositeReferencesFromRefSet(column, weakRefs, axis, out ref1, out ref2))
                {
                    return true;
                }
            }
            catch
            {
            }

            List<Face> faces = GetFaceListFromColumnSolid(opt, column, dir);
            if (faces.Count >= 2)
            {
                ref1 = faces[0]?.Reference;
                ref2 = faces[1]?.Reference;
                return ref1 != null && ref2 != null;
            }

            return false;
        }

        private static bool TryGetOppositeReferencesFromRefSet(
            FamilyInstance column,
            IList<Reference> refs,
            XYZ axis,
            out Reference refMin,
            out Reference refMax)
        {
            refMin = null;
            refMax = null;

            if (column == null || refs == null || refs.Count < 2 || axis == null)
            {
                return false;
            }

            double minProj = double.MaxValue;
            double maxProj = double.MinValue;
            XYZ axisNorm = axis.Normalize();

            foreach (Reference reference in refs)
            {
                if (reference == null)
                {
                    continue;
                }

                GeometryObject geometryObject;
                try
                {
                    geometryObject = column.GetGeometryObjectFromReference(reference);
                }
                catch
                {
                    continue;
                }

                if (!(geometryObject is Face face))
                {
                    continue;
                }

                BoundingBoxUV bbox;
                try
                {
                    bbox = face.GetBoundingBox();
                }
                catch
                {
                    continue;
                }

                if (bbox == null)
                {
                    continue;
                }

                UV centerUv = new UV(
                    (bbox.Min.U + bbox.Max.U) * 0.5,
                    (bbox.Min.V + bbox.Max.V) * 0.5);

                XYZ point;
                XYZ normal;
                try
                {
                    point = face.Evaluate(centerUv);
                    normal = face is PlanarFace planarFace
                        ? planarFace.FaceNormal
                        : face.ComputeNormal(centerUv);
                }
                catch
                {
                    continue;
                }

                XYZ normalNorm = NormalizeOrNull(normal);
                if (normalNorm == null)
                {
                    continue;
                }

                if (Math.Abs(normalNorm.DotProduct(axisNorm)) < 0.95)
                {
                    continue;
                }

                double proj = point.DotProduct(axisNorm);
                if (proj < minProj)
                {
                    minProj = proj;
                    refMin = reference;
                }

                if (proj > maxProj)
                {
                    maxProj = proj;
                    refMax = reference;
                }
            }

            return refMin != null && refMax != null && !ReferenceEquals(refMin, refMax);
        }

        private static List<Face> GetFaceListFromColumnSolid(Options opt, FamilyInstance column, int dir)
        {
            List<(Face Face, XYZ WorldNormal, XYZ WorldPoint)> candidates = new List<(Face, XYZ, XYZ)>();

            GeometryElement geomElem = column.get_Geometry(opt);
            if (geomElem == null)
            {
                return new List<Face>();
            }

            XYZ targetDir = dir == 0 ? column.HandOrientation : column.FacingOrientation;
            XYZ targetDirNeg = targetDir.Negate();

            void TryAddFace(Face face, Transform transform)
            {
                if (face == null || face.Reference == null)
                {
                    return;
                }

                UV uv = new UV(0.5, 0.5);

                XYZ normal;
                XYZ point;
                try
                {
                    normal = face.ComputeNormal(uv);
                    point = face.Evaluate(uv);
                }
                catch
                {
                    return;
                }

                XYZ worldNormal = transform != null ? transform.OfVector(normal) : normal;
                XYZ worldPoint = transform != null ? transform.OfPoint(point) : point;

                if (worldNormal.IsAlmostEqualTo(targetDir) || worldNormal.IsAlmostEqualTo(targetDirNeg))
                {
                    candidates.Add((face, worldNormal, worldPoint));
                }
            }

            void ProcessGeometryElement(GeometryElement geometry, Transform transform)
            {
                if (geometry == null)
                {
                    return;
                }

                foreach (GeometryObject obj in geometry)
                {
                    if (obj is Solid solid)
                    {
                        if (solid.Faces.Size == 0 || solid.Edges.Size == 0)
                        {
                            continue;
                        }

                        foreach (Face face in solid.Faces)
                        {
                            TryAddFace(face, transform);
                        }
                    }
                    else if (obj is GeometryInstance gi)
                    {
                        GeometryElement symbolGeometry = null;
                        try
                        {
                            symbolGeometry = gi.GetSymbolGeometry();
                        }
                        catch
                        {
                        }

                        if (symbolGeometry != null)
                        {
                            Transform nestedTransform = transform == null ? gi.Transform : transform.Multiply(gi.Transform);
                            ProcessGeometryElement(symbolGeometry, nestedTransform);
                        }

                        GeometryElement instanceGeometry = null;
                        try
                        {
                            instanceGeometry = gi.GetInstanceGeometry();
                        }
                        catch
                        {
                        }

                        if (instanceGeometry != null)
                        {
                            Transform nestedTransform = transform == null ? gi.Transform : transform.Multiply(gi.Transform);
                            ProcessGeometryElement(instanceGeometry, nestedTransform);
                        }
                    }
                }
            }

            ProcessGeometryElement(geomElem, Transform.Identity);

            if (candidates.Count < 2)
            {
                return new List<Face>();
            }

            XYZ axis = dir == 0 ? column.HandOrientation : column.FacingOrientation;

            Face minFace = null;
            Face maxFace = null;
            double minProj = double.MaxValue;
            double maxProj = double.MinValue;

            foreach (var item in candidates)
            {
                double proj = item.WorldPoint.DotProduct(axis);

                if (proj < minProj)
                {
                    minProj = proj;
                    minFace = item.Face;
                }

                if (proj > maxProj)
                {
                    maxProj = proj;
                    maxFace = item.Face;
                }
            }

            List<Face> result = new List<Face>();
            if (minFace?.Reference != null)
            {
                result.Add(minFace);
            }

            if (maxFace?.Reference != null && maxFace != minFace)
            {
                result.Add(maxFace);
            }

            return result;
        }

        private static ReferenceArray CreateSortedReferenceArrayByProjection(
            IEnumerable<(Reference Ref, double Projection)> references)
        {
            ReferenceArray result = new ReferenceArray();
            if (references == null)
            {
                return result;
            }

            foreach (var item in references
                .Where(x => x.Ref != null)
                .OrderBy(x => x.Projection))
            {
                result.Append(item.Ref);
            }

            return result;
        }

        private static void TryLockDimension(
            Dimension dimension)
        {
            if (dimension == null)
            {
                return;
            }

            try
            {
                if (!dimension.IsLocked)
                {
                    dimension.IsLocked = true;
                }
            }
            catch
            {
            }
        }

        private static ReferenceArray CreateSortedReferenceArrayAlongAxis(
            Document doc,
            IEnumerable<Reference> references,
            XYZ axis)
        {
            ReferenceArray result = new ReferenceArray();
            XYZ axisNorm = NormalizeOrNull(axis);

            if (axisNorm == null)
            {
                return result;
            }

            List<(Reference Ref, double Projection)> items = new List<(Reference, double)>();

            foreach (Reference reference in references)
            {
                if (reference == null)
                {
                    continue;
                }

                Element element = doc.GetElement(reference);
                if (element == null)
                {
                    continue;
                }

                GeometryObject geometryObject = null;
                try
                {
                    geometryObject = element.GetGeometryObjectFromReference(reference);
                }
                catch
                {
                }

                if (geometryObject is Face face)
                {
                    BoundingBoxUV bbox = face.GetBoundingBox();
                    if (bbox != null)
                    {
                        UV centerUv = new UV(
                            (bbox.Min.U + bbox.Max.U) * 0.5,
                            (bbox.Min.V + bbox.Max.V) * 0.5);

                        XYZ p = face.Evaluate(centerUv);
                        items.Add((reference, p.DotProduct(axisNorm)));
                    }
                }
                else if (geometryObject is Line line)
                {
                    XYZ p = line.Evaluate(0.5, true);
                    items.Add((reference, p.DotProduct(axisNorm)));
                }
                else
                {
                    BoundingBoxXYZ bb = element.get_BoundingBox(null);
                    if (bb != null)
                    {
                        XYZ p = (bb.Min + bb.Max) * 0.5;
                        items.Add((reference, p.DotProduct(axisNorm)));
                    }
                }
            }

            foreach (var item in items.OrderBy(x => x.Projection))
            {
                result.Append(item.Ref);
            }

            return result;
        }

        private static double GetReferenceProjection(Document doc, Reference reference, XYZ axis)
        {
            XYZ axisNorm = NormalizeOrNull(axis);
            if (reference == null || axisNorm == null)
            {
                return 0;
            }

            Element element = doc.GetElement(reference);
            if (element == null)
            {
                return 0;
            }

            GeometryObject geometryObject = null;
            try
            {
                geometryObject = element.GetGeometryObjectFromReference(reference);
            }
            catch
            {
            }

            if (geometryObject is Face face)
            {
                BoundingBoxUV bbox = face.GetBoundingBox();
                if (bbox != null)
                {
                    UV centerUv = new UV(
                        (bbox.Min.U + bbox.Max.U) * 0.5,
                        (bbox.Min.V + bbox.Max.V) * 0.5);

                    return face.Evaluate(centerUv).DotProduct(axisNorm);
                }
            }
            else if (geometryObject is Line line)
            {
                return line.Evaluate(0.5, true).DotProduct(axisNorm);
            }

            BoundingBoxXYZ bb = element.get_BoundingBox(null);
            if (bb != null)
            {
                return ((bb.Min + bb.Max) * 0.5).DotProduct(axisNorm);
            }

            return 0;
        }

        private static bool TryGetColumnDimensions(
            FamilyInstance column,
            Options opt,
            XYZ handOrientation,
            XYZ facingOrientation,
            out double columnWidth,
            out double columnHeight)
        {
            columnWidth = 0;
            columnHeight = 0;

            GeometryElement geometry = column.get_Geometry(opt);
            if (geometry == null)
            {
                return false;
            }

            List<XYZ> points = new List<XYZ>();
            CollectSolidPoints(geometry, points);
            if (points.Count < 2)
            {
                return false;
            }

            XYZ hand = handOrientation.Normalize();
            XYZ facing = facingOrientation.Normalize();

            double minHand = double.MaxValue;
            double maxHand = double.MinValue;
            double minFacing = double.MaxValue;
            double maxFacing = double.MinValue;

            foreach (XYZ point in points)
            {
                double handValue = point.DotProduct(hand);
                double facingValue = point.DotProduct(facing);

                if (handValue < minHand) minHand = handValue;
                if (handValue > maxHand) maxHand = handValue;
                if (facingValue < minFacing) minFacing = facingValue;
                if (facingValue > maxFacing) maxFacing = facingValue;
            }

            columnWidth = maxHand - minHand;
            columnHeight = maxFacing - minFacing;

            return columnWidth > Eps && columnHeight > Eps;
        }

        private static void CollectSolidPoints(GeometryElement geometry, List<XYZ> points)
        {
            foreach (GeometryObject obj in geometry)
            {
                if (obj is Solid solid)
                {
                    if (solid.Faces.Size == 0 || solid.Edges.Size == 0 || solid.Volume == 0)
                    {
                        continue;
                    }

                    foreach (Edge edge in solid.Edges)
                    {
                        IList<XYZ> edgePoints = edge.Tessellate();
                        foreach (XYZ point in edgePoints)
                        {
                            points.Add(point);
                        }
                    }
                }
                else if (obj is GeometryInstance instance)
                {
                    GeometryElement instanceGeometry = instance.GetInstanceGeometry();
                    if (instanceGeometry != null)
                    {
                        CollectSolidPoints(instanceGeometry, points);
                    }
                }
            }
        }

        private static void AdjustSingleDimensionText(
            Document doc,
            ElementId viewId,
            ElementId textTypeId,
            Dimension dimension,
            double dimensionLineLength,
            XYZ axis,
            int scale)
        {
            if (dimension == null)
            {
                return;
            }

            string text = dimension.get_Parameter(BuiltInParameter.DIM_VALUE_LENGTH)?.AsValueString();
            if (string.IsNullOrWhiteSpace(text))
            {
                return;
            }

            double textWidth = MeasureTextWidth(doc, viewId, textTypeId, text, scale);
            if (textWidth <= dimensionLineLength)
            {
                return;
            }

            using (Transaction t = new Transaction(doc, "Скорректировать текст размера"))
            {
                t.Start();

                XYZ direction = axis.Normalize();
                dimension.TextPosition =
                    dimension.TextPosition
                    + direction * (dimensionLineLength / 2.0)
                    + direction * (textWidth / 2.0)
                    + direction * (2.0 / 304.8 * scale);

                t.Commit();
            }
        }

        private static void MoveSegmentTexts(
            Document doc,
            ElementId viewId,
            ElementId textTypeId,
            Dimension dimension,
            XYZ axis,
            int scale)
        {
            if (dimension == null || dimension.Segments == null || dimension.Segments.Size == 0)
            {
                return;
            }

            XYZ directionBase = axis.Normalize();
            List<double> originProjections = new List<double>();
            foreach (DimensionSegment seg in dimension.Segments)
            {
                if (seg?.Origin != null)
                {
                    originProjections.Add(seg.Origin.DotProduct(directionBase));
                }
            }

            double centerProjection = originProjections.Count > 0
                ? originProjections.Average()
                : 0.0;

            foreach (DimensionSegment segment in dimension.Segments)
            {
                double segmentValue = segment?.Value ?? 0.0;
                if (segmentValue <= Eps)
                {
                    continue;
                }

                string text = segment.ValueString;
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                double textWidth = MeasureTextWidth(doc, viewId, textTypeId, text, scale);
                // Добавляем технологический зазор, иначе при почти равных значениях
                // текст визуально всё равно "слипается" со стрелками.
                double fitThreshold = textWidth + 2.0 * (2.0 / 304.8 * scale);
                if (fitThreshold <= segmentValue)
                {
                    continue;
                }

                using (Transaction t = new Transaction(doc, "Сдвиг текста сегмента"))
                {
                    t.Start();

                    if (segment.Origin == null || segment.TextPosition == null)
                    {
                        t.RollBack();
                        continue;
                    }

                    double segmentProjection = segment.Origin.DotProduct(directionBase);
                    XYZ direction = segmentProjection <= centerProjection
                        ? directionBase.Negate()
                        : directionBase;

                    segment.TextPosition =
                        segment.TextPosition
                        + direction * (segmentValue / 2.0)
                        + direction * (textWidth / 2.0)
                        + direction * (2.0 / 304.8 * scale);

                    t.Commit();
                }
            }
        }

        private static double MeasureTextWidth(
            Document doc,
            ElementId viewId,
            ElementId textTypeId,
            string text,
            int scale)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return 0;
            }

            using (Transaction t = new Transaction(doc, "Измерить ширину текста"))
            {
                t.Start();

                TextNote tempText = TextNote.Create(doc, viewId, XYZ.Zero, text, textTypeId);
                double width = tempText.Width * scale;
                doc.Delete(tempText.Id);

                t.Commit();
                return width;
            }
        }

        private static bool TryGetOrCreateTempTextType(
            Document doc,
            DimensionType dimensionType,
            out TextNoteType textNoteType,
            out bool createdNew)
        {
            textNoteType = null;
            createdNew = false;

            using (Transaction t = new Transaction(doc, "Создать временный тип текста"))
            {
                t.Start();

                textNoteType = new FilteredElementCollector(doc)
                    .OfClass(typeof(TextNoteType))
                    .WhereElementIsElementType()
                    .Cast<TextNoteType>()
                    .FirstOrDefault(x => x.Name == "tmpTT");

                if (textNoteType == null)
                {
                    TextNoteType baseType = new FilteredElementCollector(doc)
                        .OfClass(typeof(TextNoteType))
                        .WhereElementIsElementType()
                        .Cast<TextNoteType>()
                        .FirstOrDefault();

                    if (baseType == null)
                    {
                        t.RollBack();
                        return false;
                    }

                    textNoteType = baseType.Duplicate("tmpTT") as TextNoteType;
                    if (textNoteType == null)
                    {
                        t.RollBack();
                        return false;
                    }

                    createdNew = true;

                    string fontName = dimensionType.get_Parameter(BuiltInParameter.TEXT_FONT)?.AsString();
                    double textSizeMm = (dimensionType.get_Parameter(BuiltInParameter.TEXT_SIZE)?.AsDouble() ?? 0.0) * 304.8;
                    int styleBold = dimensionType.get_Parameter(BuiltInParameter.TEXT_STYLE_BOLD)?.AsInteger() ?? 0;
                    int styleItalic = dimensionType.get_Parameter(BuiltInParameter.TEXT_STYLE_ITALIC)?.AsInteger() ?? 0;
                    int styleUnderline = dimensionType.get_Parameter(BuiltInParameter.TEXT_STYLE_UNDERLINE)?.AsInteger() ?? 0;
                    double widthScale = dimensionType.get_Parameter(BuiltInParameter.TEXT_WIDTH_SCALE)?.AsDouble() ?? 1.0;

                    if (!string.IsNullOrWhiteSpace(fontName))
                    {
                        textNoteType.get_Parameter(BuiltInParameter.TEXT_FONT)?.Set(fontName);
                    }

#if R2019 || R2020 || R2021
                    textNoteType.get_Parameter(BuiltInParameter.TEXT_SIZE)?.Set(
                        UnitUtils.ConvertToInternalUnits(textSizeMm, DisplayUnitType.DUT_MILLIMETERS));
#else
                    textNoteType.get_Parameter(BuiltInParameter.TEXT_SIZE)?.Set(
                        UnitUtils.ConvertToInternalUnits(textSizeMm, UnitTypeId.Millimeters));
#endif
                    textNoteType.get_Parameter(BuiltInParameter.TEXT_STYLE_BOLD)?.Set(styleBold);
                    textNoteType.get_Parameter(BuiltInParameter.TEXT_STYLE_ITALIC)?.Set(styleItalic);
                    textNoteType.get_Parameter(BuiltInParameter.TEXT_STYLE_UNDERLINE)?.Set(styleUnderline);
                    textNoteType.get_Parameter(BuiltInParameter.TEXT_WIDTH_SCALE)?.Set(widthScale);
                }

                t.Commit();
                return textNoteType != null;
            }
        }

        private static int GetViewScale(View view)
        {
            Parameter scaleParameter = view.get_Parameter(BuiltInParameter.VIEW_SCALE_PULLDOWN_METRIC);
            int scale = scaleParameter != null ? scaleParameter.AsInteger() : 100;
            return scale > 0 ? scale : 100;
        }

        private static XYZ NormalizeOrNull(XYZ vector)
        {
            if (vector == null)
            {
                return null;
            }

            double length = vector.GetLength();
            if (length < Eps)
            {
                return null;
            }

            return vector / length;
        }

        private static bool TryParseOffsetInMillimeters(string rawValue, out double valueInFeet)
        {
            valueInFeet = 0;

            if (string.IsNullOrWhiteSpace(rawValue))
            {
                return false;
            }

            bool parsed =
                double.TryParse(rawValue, NumberStyles.Float, CultureInfo.CurrentCulture, out double valueMm) ||
                double.TryParse(rawValue, NumberStyles.Float, CultureInfo.InvariantCulture, out valueMm);

            if (!parsed || valueMm < 0)
            {
                return false;
            }

            valueInFeet = valueMm / 304.8;
            return true;
        }

        private static List<FamilyInstance> GetColumns(
            Document doc,
            Selection sel,
            View activeView,
            ColumnDimensionerWPF window)
        {
            if (window.DimensionColumnsButtonName == "radioButton_VisibleInView")
            {
                List<FamilyInstance> visibleColumns = new FilteredElementCollector(doc, activeView.Id)
                    .OfCategory(BuiltInCategory.OST_StructuralColumns)
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>()
                    .ToList();
                return visibleColumns;
            }

            List<FamilyInstance> selectedColumns = GetColumnsFromCurrentSelection(doc, sel);
            if (selectedColumns.Count > 0)
            {
                return selectedColumns;
            }

            ColumnSelectionFilter filter = new ColumnSelectionFilter();
            IList<Reference> pickedRefs;

            try
            {
                pickedRefs = sel.PickObjects(ObjectType.Element, filter, "Выберите несущие колонны!");
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return new List<FamilyInstance>();
            }

            List<FamilyInstance> pickedColumns = pickedRefs
                .Select(r => doc.GetElement(r))
                .OfType<FamilyInstance>()
                .ToList();
            return pickedColumns;
        }

        private static List<FamilyInstance> GetColumnsFromCurrentSelection(Document doc, Selection sel)
        {
            ICollection<ElementId> selectedIds = sel.GetElementIds();
            List<FamilyInstance> result = new List<FamilyInstance>();

            foreach (ElementId id in selectedIds)
            {
                if (doc.GetElement(id) is FamilyInstance fi
                    && fi.Category != null
                    && fi.Category.Id == new ElementId(BuiltInCategory.OST_StructuralColumns))
                {
                    result.Add(fi);
                }
            }
            return result;
        }

        private static string IdToText(ElementId id)
        {
            return id?.ToString() ?? "<null>";
        }

        private static async Task GetPluginStartInfo()
        {
            Assembly thisAssembly = Assembly.GetExecutingAssembly();
            string assemblyName = "ColumnDimensioner";
            string assemblyNameRus = "Образмерить колонны";
            string assemblyFolderPath = Path.GetDirectoryName(thisAssembly.Location);

            int lastBackslashIndex = assemblyFolderPath.LastIndexOf("\\");
            string dllPath = assemblyFolderPath.Substring(0, lastBackslashIndex + 1) + "PluginInfoCollector\\PluginInfoCollector.dll";

            Assembly assembly = Assembly.LoadFrom(dllPath);
            Type type = assembly.GetType("PluginInfoCollector.InfoCollector");

            if (type != null)
            {
                object instance = Activator.CreateInstance(type);
                MethodInfo method = type.GetMethod("CollectPluginUsageAsync");

                if (method != null)
                {
                    Task task = (Task)method.Invoke(instance, new object[] { assemblyName, assemblyNameRus });
                    await task;
                }
            }
        }
    }
}
