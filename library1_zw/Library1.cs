using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using zzr = ZwSoft.ZwCAD.Runtime;
using zza = ZwSoft.ZwCAD.ApplicationServices;
using zzd = ZwSoft.ZwCAD.DatabaseServices;
using zze = ZwSoft.ZwCAD.EditorInput;
using zzg = ZwSoft.ZwCAD.Geometry;
using zzc = ZwSoft.ZwCAD.Colors;
using System.Globalization;
using System.Windows;

namespace library1_zw
{
    class Library1
    {
       

        /// <summary>
        /// Sprawdza czy jesteśmy w przestrzenii modelu
        /// </summary>
        /// <returns></returns>        
        public static bool ItisModel()
        {
            if (zza.Application.DocumentManager.MdiActiveDocument.Database.TileMode)
                return true;
            else
                return false;
        }
        
        /// <summary>
        /// Sprawdza czy w przestrzeni papieru jesteśmy  poza oknem true lub w oknie false
        /// na podstawie https://spiderinnet1.typepad.com/blog/2014/05/autocad-net-detect-current-space-model-or-paper-and-viewport.html
        /// </summary>
        /// <returns></returns>
        public static bool IsInLayoutPaper()
        {
            zza.Document doc = zza.Application.DocumentManager.MdiActiveDocument;
            zzd.Database db = doc.Database;
            zze.Editor ed = doc.Editor;

            if (db.TileMode)
                return false;
            else
            {
                if (db.PaperSpaceVportId == zzd.ObjectId.Null)
                    return false;
                else if (ed.CurrentViewportObjectId == zzd.ObjectId.Null)
                    return false;
                else if (ed.CurrentViewportObjectId == db.PaperSpaceVportId)
                    return true;
                else
                    return false;
            }
        }

        /// <summary>
        /// Wstawia warstwę o podanej nazwie, kolorze i czyni ją plotowaną lub nie.
        /// </summary>
        /// <param name="layername"></param>
        /// <param name="color"></param>
        /// <param name="isplotable"></param>
        public static void WstawWarstwe(string layername, short color, bool isplotable)
        {
            zza.Document doc = zza.Application.DocumentManager.MdiActiveDocument;
            zzd.Database db = doc.Database;
            using (zzd.Transaction tr = db.TransactionManager.StartTransaction())
            {
                
                zzd.LayerTable lt = (zzd.LayerTable)tr.GetObject(db.LayerTableId, zzd.OpenMode.ForRead);
                if (lt.Has(layername) == false)
                {
                    // Tworzymy nową wartwę papier
                    zzd.LayerTableRecord nowawarstwa = new zzd.LayerTableRecord();
                    //nadajemy jej wlasciwosci
                    nowawarstwa.Name = layername;
                    nowawarstwa.IsPlottable = isplotable;
                    nowawarstwa.Color = zzc.Color.FromColorIndex(zzc.ColorMethod.ByAci, color);
                    //dadaj nowarstwa do tabeli
                    lt.UpgradeOpen();
                    zzd.ObjectId warstwa = lt.Add(nowawarstwa);
                    tr.AddNewlyCreatedDBObject(nowawarstwa, true);
                }

                tr.Commit();
            }
        }

        /// <summary>
        /// Pobier lub wstawia styltekstu PI_DIMENSIONTEXT
        /// </summary>
        /// <returns></returns>
        public static zzd.ObjectId GetSetPI_DIMENSIONTEXT()
        {
            string name = "PI_DIMENSIONTEXT";         

            zzd.ObjectId dimstyleIDE;
            zza.Document doc = zza.Application.DocumentManager.MdiActiveDocument;
            zzd.Database db = doc.Database;
            using (zzd.Transaction tr = db.TransactionManager.StartTransaction())
            {
                zzd.TextStyleTable tst = (zzd.TextStyleTable)tr.GetObject(db.TextStyleTableId, zzd.OpenMode.ForWrite);

                if (!tst.Has(name))
                {
                    tst.UpgradeOpen();
                    zzd.TextStyleTableRecord newRecord = new zzd.TextStyleTableRecord();
                    newRecord.Name = name;
                    newRecord.FileName = "simplex.shx";
                    newRecord.XScale = 0.65; // Width factor
                    tst.Add(newRecord);
                    tr.AddNewlyCreatedDBObject(newRecord, true);
                    dimstyleIDE = tst[name];
                }
                else
                {
                    dimstyleIDE = tst[name];
                }

                tr.Commit();
           }

            return dimstyleIDE;
        }
        

        /// <summary>
        /// Wstawia kotę wysokościową używaną do oznaczania poziomu głównej konstrukcji nośnej
        /// </summary>
        public static void Kota_Kon()
        {
            for (; ; )
            {
                //wstawiamy warstwę jeśli nie istnieje
                string layername = "PI_KOTA_WYS";
                string startpromt = "\nWskaż punkt wysokościowy: ";
                string endprompt = "\nWskaż dół/góra: ";
                // wstawiamy jeśli nie istnieje warstwę
                WstawWarstwe(layername, 1, true);
                // wstawiamy jeśli nie istnieje styl wymiarowania
                zzd.ObjectId stylwymiarow = GetSetPI_DIMENSIONTEXT();                  
                //pobieramy bazę danych aktualnego rysunku
                zza.Document doc = zza.Application.DocumentManager.MdiActiveDocument;
                zzd.Database db = doc.Database;
                zze.Editor ed = doc.Editor;
                //zapiszemy aktualną zmienną osmode
                int osmode = Convert.ToInt16(zza.Application.GetSystemVariable("OSMODE"));

                bool jestmodel = ItisModel();
                bool jestviewport = IsInLayoutPaper();

                zze.PromptPointResult pPtRes;
                zze.PromptPointOptions pPtOpts = new zze.PromptPointOptions("");

                //przełącz zmienną osmode na 512
                zza.Application.SetSystemVariable("Osmode", 512);

                // zapytaj o srodek
                pPtOpts.Message = startpromt;
                pPtRes = doc.Editor.GetPoint(pPtOpts);
                zzg.Point3d ptStart = pPtRes.Value;

                // Exit if the user presses ESC or cancels the command
                if (pPtRes.Status == zze.PromptStatus.Cancel)
                {
                    zza.Application.SetSystemVariable("Osmode", osmode);
                    return;
                }

                // zapytaj o koniec przekroju
                pPtOpts.Message = endprompt ;
                pPtOpts.UseBasePoint = true;
                pPtOpts.BasePoint = ptStart;
                pPtRes = doc.Editor.GetPoint(pPtOpts);
                zzg.Point3d ptEnd = pPtRes.Value;

                if (pPtRes.Status == zze.PromptStatus.Cancel)
                {
                    zza.Application.SetSystemVariable("Osmode", osmode);
                    return;
                }

                //conversja
                zzg.Matrix3d ucs = ed.CurrentUserCoordinateSystem;
                zzg.CoordinateSystem3d cs = ucs.CoordinateSystem3d;

                // Transform from UCS to WCS
                zzg.Matrix3d mat = zzg.Matrix3d.AlignCoordinateSystem(zzg.Point3d.Origin, zzg.Vector3d.XAxis, zzg.Vector3d.YAxis, zzg.Vector3d.ZAxis,
                    cs.Origin, cs.Xaxis, cs.Yaxis, cs.Zaxis);

                //pobieramy wartość zmiennej dimscale
                double dimscala = Convert.ToDouble(zza.Application.GetSystemVariable("DIMSCALE"));
                //pobieramy wartość zmiennej  insunits
                int insunits = Convert.ToUInt16(zza.Application.GetSystemVariable("INSUNITS"));

                using (zzd.Transaction tr = db.TransactionManager.StartTransaction())
                {
                    //utworzenie grupy
                    zzd.DBDictionary groupDic = (zzd.DBDictionary)tr.GetObject(db.GroupDictionaryId, zzd.OpenMode.ForWrite);
                    zzd.Group anonyGroup = new zzd.Group();
                    groupDic.SetAt("*", anonyGroup);

                    zzd.BlockTable bt;
                    zzd.BlockTableRecord btr;

                    // Open Model space for write
                    bt = tr.GetObject(db.BlockTableId,
                                                    zzd.OpenMode.ForRead) as zzd.BlockTable;
                    if (jestmodel == false & jestviewport == true)
                    {
                        btr = tr.GetObject(bt[zzd.BlockTableRecord.PaperSpace],
                                   zzd.OpenMode.ForWrite) as zzd.BlockTableRecord;
                        dimscala = 1;
                    }
                    else
                    {
                        btr = tr.GetObject(bt[zzd.BlockTableRecord.ModelSpace],zzd.OpenMode.ForWrite) as zzd.BlockTableRecord;
                    }
                    double tekst_dlugosc;

                    //rysujemy tekst
                    using (zzd.DBText acText = new zzd.DBText())
                    {
                        double poziom = ptStart.Y;
                        if (insunits == 4) poziom = poziom / 1000.0;
                        if (insunits == 5) poziom = poziom / 100.0;

                        if (ptStart.Y >= -0.001 & ptStart.Y <= 0.001)
                        {
                            acText.TextString = "%%p " + poziom.ToString("0.000", CultureInfo.GetCultureInfo("pl-PL"));
                        }
                        else
                        {
                            if (ptStart.Y > 0.001)
                            {
                                acText.TextString = "+ " + poziom.ToString("0.000", CultureInfo.GetCultureInfo("pl-PL"));
                            }
                            else
                            {
                                acText.TextString = "- " + Math.Abs(poziom).ToString("0.000", CultureInfo.GetCultureInfo("pl-PL"));
                            }
                        }

                        acText.TextStyleId = stylwymiarow;
                        acText.HorizontalMode = zzd.TextHorizontalMode.TextCenter;
                        acText.VerticalMode = zzd.TextVerticalMode.TextTop;
                        if (ptStart.Y < ptEnd.Y)
                        {
                            acText.AlignmentPoint = new zzg.Point3d(ptStart.X - 2 * dimscala, ptStart.Y + dimscala * 5, ptStart.Z);
                        }
                        else
                        {
                            acText.AlignmentPoint = new zzg.Point3d(ptStart.X - 2 * dimscala, ptStart.Y - dimscala * 5, ptStart.Z);
                        }
                        acText.Height = 2 * dimscala;
                        acText.WidthFactor = 0.65;
                        acText.Layer = layername;
                        acText.ColorIndex = 2;


                        btr.AppendEntity(acText);
                        tr.AddNewlyCreatedDBObject(acText, true);
                        anonyGroup.Append(acText.ObjectId);

                        zzg.Point3d ptMax2 = acText.GeometricExtents.MaxPoint;
                        zzg.Point3d ptMin2 = acText.GeometricExtents.MinPoint;

                        tekst_dlugosc = ptMax2.X - ptMin2.X;
                        acText.Erase();
                    }
                    //rysujemy tekst
                    using (zzd.DBText acText = new zzd.DBText())
                    {

                        double poziom = ptStart.Y;
                        if (insunits == 4) poziom = poziom / 1000.0;
                        if (insunits == 5) poziom = poziom / 100.0;

                        if (ptStart.Y >= -0.001 & ptStart.Y <= 0.001)
                        {

                            acText.TextString = "%%p " + poziom.ToString("0.000", CultureInfo.GetCultureInfo("pl-PL"));
                            // acText.TextString = "%%p" + poziom.ToString();
                        }
                        else
                        {
                            if (ptStart.Y > 0.001)
                            {
                                acText.TextString = "+ " + poziom.ToString("0.000", CultureInfo.GetCultureInfo("pl-PL"));
                                //acText.TextString = "+" + poziom.ToString();
                            }
                            else
                            {
                                acText.TextString = "- " + Math.Abs(poziom).ToString("0.000", CultureInfo.GetCultureInfo("pl-PL"));
                                //acText.TextString = poziom.ToString();
                            }
                        }
                        acText.TextStyleId = stylwymiarow;
                        acText.HorizontalMode = zzd.TextHorizontalMode.TextLeft;
                        acText.VerticalMode = zzd.TextVerticalMode.TextVerticalMid;
                        if (ptStart.Y < ptEnd.Y)
                        {
                            acText.AlignmentPoint = new zzg.Point3d(ptStart.X - 2 * dimscala, ptStart.Y + dimscala * 5, ptStart.Z);
                        }
                        else
                        {
                            acText.AlignmentPoint = new zzg.Point3d(ptStart.X - 2 * dimscala, ptStart.Y - dimscala * 5, ptStart.Z);
                        }
                        acText.Height = 2 * dimscala;
                        acText.WidthFactor = 0.65;
                        acText.Layer = layername;
                        acText.ColorIndex = 2;
                        acText.TransformBy(mat);

                        btr.AppendEntity(acText);
                        tr.AddNewlyCreatedDBObject(acText, true);
                        anonyGroup.Append(acText.ObjectId);
                    }

                    //rysujemy polilinie trójkąta               
                    using (zzd.Polyline acPoly = new zzd.Polyline())
                    {

                        if (ptStart.Y < ptEnd.Y)
                        {
                            acPoly.AddVertexAt(0, new zzg.Point2d(ptStart.X, ptStart.Y + 7 * dimscala), 0, 0, 0);
                            acPoly.AddVertexAt(1, new zzg.Point2d(ptStart.X, ptStart.Y), 0, 0, 0);
                            acPoly.AddVertexAt(2, new zzg.Point2d(ptStart.X - 2 * dimscala, ptStart.Y + 3 * dimscala), 0, 0, 0);
                            acPoly.AddVertexAt(3, new zzg.Point2d(ptStart.X - 2 * dimscala + tekst_dlugosc, ptStart.Y + 3 * dimscala), 0, 0, 0);

                            acPoly.Layer = layername;
                            acPoly.ColorIndex = 1;
                            acPoly.TransformBy(mat);

                            btr.AppendEntity(acPoly);
                            tr.AddNewlyCreatedDBObject(acPoly, true);
                            anonyGroup.Append(acPoly.ObjectId);
                        }
                        else
                        {
                            acPoly.AddVertexAt(0, new zzg.Point2d(ptStart.X, ptStart.Y - 7 * dimscala), 0, 0, 0);
                            acPoly.AddVertexAt(1, new zzg.Point2d(ptStart.X, ptStart.Y), 0, 0, 0);
                            acPoly.AddVertexAt(2, new zzg.Point2d(ptStart.X - 2 * dimscala, ptStart.Y - 3 * dimscala), 0, 0, 0);
                            acPoly.AddVertexAt(3, new zzg.Point2d(ptStart.X - 2 * dimscala + tekst_dlugosc, ptStart.Y - 3 * dimscala), 0, 0, 0);

                            acPoly.Layer = layername;
                            acPoly.ColorIndex = 1;
                            acPoly.TransformBy(mat);

                            btr.AppendEntity(acPoly);
                            tr.AddNewlyCreatedDBObject(acPoly, true);
                            anonyGroup.Append(acPoly.ObjectId);
                        }
                    }

                    //rysujemy polilinie trójkąta               
                    using (zzd.Polyline acPoly = new zzd.Polyline())
                    {

                        if (ptStart.Y < ptEnd.Y)
                        {
                            acPoly.AddVertexAt(0, new zzg.Point2d(ptStart.X, ptStart.Y + 3 * dimscala), 0, 0, 0);
                            acPoly.AddVertexAt(1, new zzg.Point2d(ptStart.X, ptStart.Y), 0, 0, 0);
                            acPoly.AddVertexAt(2, new zzg.Point2d(ptStart.X - 2 * dimscala, ptStart.Y + 3 * dimscala), 0, 0, 0);
                            acPoly.Closed = true;

                            acPoly.Layer = layername;
                            acPoly.ColorIndex = 1;
                            acPoly.TransformBy(mat);

                            btr.AppendEntity(acPoly);
                            tr.AddNewlyCreatedDBObject(acPoly, true);

                            zzd.ObjectIdCollection acObjIdColl = new zzd.ObjectIdCollection();
                            acObjIdColl.Add(acPoly.ObjectId);

                            // Create the hatch object and append it to the block table record
                            using (zzd.Hatch acHatch = new zzd.Hatch())
                            {
                                btr.AppendEntity(acHatch);
                                tr.AddNewlyCreatedDBObject(acHatch, true);
                                anonyGroup.Append(acHatch.ObjectId);

                                // Set the properties of the hatch object
                                // Associative must be set after the hatch object is appended to the 
                                // block table record and before AppendLoop
                                acHatch.Layer = layername;
                                acHatch.ColorIndex = 2;
                                acHatch.SetHatchPattern(zzd.HatchPatternType.PreDefined, "SOLID");
                                acHatch.Associative = true;
                                acHatch.AppendLoop(zzd.HatchLoopTypes.Default, acObjIdColl);
                                acHatch.EvaluateHatch(true);
                            }
                            acPoly.Erase();

                        }
                        else
                        {
                            acPoly.AddVertexAt(0, new zzg.Point2d(ptStart.X, ptStart.Y - 3 * dimscala), 0, 0, 0);
                            acPoly.AddVertexAt(1, new zzg.Point2d(ptStart.X, ptStart.Y), 0, 0, 0);
                            acPoly.AddVertexAt(2, new zzg.Point2d(ptStart.X - 2 * dimscala, ptStart.Y - 3 * dimscala), 0, 0, 0);
                            acPoly.Closed = true;

                            acPoly.Layer = layername;
                            acPoly.ColorIndex = 1;
                            acPoly.TransformBy(mat);

                            btr.AppendEntity(acPoly);
                            tr.AddNewlyCreatedDBObject(acPoly, true);

                            zzd.ObjectIdCollection acObjIdColl = new zzd.ObjectIdCollection();
                            acObjIdColl.Add(acPoly.ObjectId);

                            // Create the hatch object and append it to the block table record
                            using (zzd.Hatch acHatch = new zzd.Hatch())
                            {
                                btr.AppendEntity(acHatch);
                                tr.AddNewlyCreatedDBObject(acHatch, true);
                                anonyGroup.Append(acHatch.ObjectId);

                                // Set the properties of the hatch object
                                // Associative must be set after the hatch object is appended to the 
                                // block table record and before AppendLoop
                                acHatch.Layer = layername;
                                acHatch.ColorIndex = 2;
                                acHatch.SetHatchPattern(zzd.HatchPatternType.PreDefined, "SOLID");
                                acHatch.Associative = true;
                                acHatch.AppendLoop(zzd.HatchLoopTypes.Default, acObjIdColl);
                                acHatch.EvaluateHatch(true);
                            }
                            acPoly.Erase();
                        }
                    }
                    tr.AddNewlyCreatedDBObject(anonyGroup, true);
                    // Commit the changes and dispose of the transaction
                    tr.Commit();
                }
                
                if (pPtRes.Status == zze.PromptStatus.Cancel)
                {   //przełącz zmienną osmode po staremu
                    zza.Application.SetSystemVariable("Osmode", osmode);
                    return;
                }
            }                        

        }

        /// <summary>
        /// Wstawia kotę wysokościową używaną do oznaczania poziomów wykończonych
        /// </summary>
        public static void Kota_Wyk()
        {
            for (; ;)
            {
                //wstawiamy warstwę jeśli nie istnieje
                string layername = "PI_KOTA_WYS";
                string startpromt = "\nWskaż punkt wysokościowy: ";
                string endprompt = "\nWskaż dół/góra: ";
                // wstawiamy jeśli nie istnieje warstwę
                WstawWarstwe(layername, 1, true);
                // wstawiamy jeśli nie istnieje styl wymiarowania
                zzd.ObjectId stylwymiarow = GetSetPI_DIMENSIONTEXT();

                zza.Document doc = zza.Application.DocumentManager.MdiActiveDocument;
                zzd.Database db = doc.Database;
                zze.Editor ed = doc.Editor;

                //zapiszemy aktualną zmienną osmode
                int osmode = Convert.ToInt16(zza.Application.GetSystemVariable("OSMODE"));

                bool jestmodel = ItisModel();
                bool jestviewport = IsInLayoutPaper();

                zze.PromptPointResult pPtRes;
                zze.PromptPointOptions pPtOpts = new zze.PromptPointOptions("");

                //przełącz zmienną osmode na 512
                zza.Application.SetSystemVariable("Osmode", 512);

                // zapytaj o srodek
                pPtOpts.Message = startpromt;
                pPtRes = doc.Editor.GetPoint(pPtOpts);
                zzg.Point3d ptStart = pPtRes.Value;

                // Exit if the user presses ESC or cancels the command
                if (pPtRes.Status == zze.PromptStatus.Cancel)
                {
                    zza.Application.SetSystemVariable("Osmode", osmode);
                    return;
                }
                
                    // zapytaj o koniec przekroju
                    pPtOpts.Message = endprompt;
                    pPtOpts.UseBasePoint = true;
                    pPtOpts.BasePoint = ptStart;
                    pPtRes = doc.Editor.GetPoint(pPtOpts);
                    zzg.Point3d ptEnd = pPtRes.Value;

                    if (pPtRes.Status == zze.PromptStatus.Cancel)
                    {
                        zza.Application.SetSystemVariable("Osmode", osmode);
                        return;
                    }

                    //conversja
                    zzg.Matrix3d ucs = ed.CurrentUserCoordinateSystem;
                    zzg.CoordinateSystem3d cs = ucs.CoordinateSystem3d;

                    // Transform from UCS to WCS
                    zzg.Matrix3d mat = zzg.Matrix3d.AlignCoordinateSystem(zzg.Point3d.Origin, zzg.Vector3d.XAxis, zzg.Vector3d.YAxis, zzg.Vector3d.ZAxis,
                        cs.Origin, cs.Xaxis, cs.Yaxis, cs.Zaxis);

                    //pobieramy wartość zmiennej dimscale
                    double dimscala = Convert.ToDouble(zza.Application.GetSystemVariable("DIMSCALE"));
                    //pobieramy wartość zmiennej  insunits
                    int insunits = Convert.ToUInt16(zza.Application.GetSystemVariable("INSUNITS"));

                    using (zzd.Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        //utworzenie grupy
                        zzd.DBDictionary groupDic = (zzd.DBDictionary)tr.GetObject(db.GroupDictionaryId, zzd.OpenMode.ForWrite);
                        zzd.Group anonyGroup = new zzd.Group();
                        groupDic.SetAt("*", anonyGroup);

                        zzd.BlockTable bt;
                        zzd.BlockTableRecord btr;

                        // Open Model space for write
                        bt = tr.GetObject(db.BlockTableId,
                                                        zzd.OpenMode.ForRead) as zzd.BlockTable;
                        if (jestmodel == false & jestviewport == true)
                        {
                            btr = tr.GetObject(bt[zzd.BlockTableRecord.PaperSpace],
                                       zzd.OpenMode.ForWrite) as zzd.BlockTableRecord;
                            dimscala = 1;
                        }
                        else
                        {
                            btr = tr.GetObject(bt[zzd.BlockTableRecord.ModelSpace],
                                                       zzd.OpenMode.ForWrite) as zzd.BlockTableRecord;
                        }

                        double tekst_dlugosc;

                        //rysujemy tekst
                        using (zzd.DBText acText = new zzd.DBText())
                        {
                            double poziom = ptStart.Y;
                            if (insunits == 4) poziom = poziom / 1000.0;
                            if (insunits == 5) poziom = poziom / 100.0;



                            if (ptStart.Y >= -0.001 & ptStart.Y <= 0.001)
                            {

                                acText.TextString = "%%p " + poziom.ToString("0.000", CultureInfo.GetCultureInfo("pl-PL"));
                                // acText.TextString = "%%p" + poziom.ToString();
                            }
                            else
                            {
                                if (ptStart.Y > 0.001)
                                {
                                    acText.TextString = "+ " + poziom.ToString("0.000", CultureInfo.GetCultureInfo("pl-PL"));
                                    //acText.TextString = "+" + poziom.ToString();
                                }
                                else
                                {
                                    acText.TextString = "- " + Math.Abs(poziom).ToString("0.000", CultureInfo.GetCultureInfo("pl-PL"));
                                    //acText.TextString = poziom.ToString();
                                }
                            }

                            acText.TextStyleId = stylwymiarow;
                            acText.HorizontalMode = zzd.TextHorizontalMode.TextCenter;
                            acText.VerticalMode = zzd.TextVerticalMode.TextTop;
                            if (ptStart.Y < ptEnd.Y)
                            {
                                acText.AlignmentPoint = new zzg.Point3d(ptStart.X - 2 * dimscala, ptStart.Y + dimscala * 5, ptStart.Z);
                            }
                            else
                            {
                                acText.AlignmentPoint = new zzg.Point3d(ptStart.X - 2 * dimscala, ptStart.Y - dimscala * 5, ptStart.Z);
                            }
                            acText.Height = 2 * dimscala;
                            acText.WidthFactor = 0.65;
                            acText.Layer = layername;
                            acText.ColorIndex = 2;


                            btr.AppendEntity(acText);
                            tr.AddNewlyCreatedDBObject(acText, true);
                            anonyGroup.Append(acText.ObjectId);

                            zzg.Point3d ptMax2 = acText.GeometricExtents.MaxPoint;
                            zzg.Point3d ptMin2 = acText.GeometricExtents.MinPoint;

                            tekst_dlugosc = ptMax2.X - ptMin2.X;
                            acText.Erase();

                        }

                        //rysujemy tekst
                        using (zzd.DBText acText = new zzd.DBText())
                        {

                            double poziom = ptStart.Y;
                            if (insunits == 4) poziom = poziom / 1000.0;
                            if (insunits == 5) poziom = poziom / 100.0;

                            if (ptStart.Y >= -0.001 & ptStart.Y <= 0.001)
                            {

                                acText.TextString = "%%p " + poziom.ToString("0.000", CultureInfo.GetCultureInfo("pl-PL"));
                                // acText.TextString = "%%p" + poziom.ToString();
                            }
                            else
                            {
                                if (ptStart.Y > 0.001)
                                {
                                    acText.TextString = "+ " + poziom.ToString("0.000", CultureInfo.GetCultureInfo("pl-PL"));
                                    //acText.TextString = "+" + poziom.ToString();
                                }
                                else
                                {
                                    acText.TextString = "- " + Math.Abs(poziom).ToString("0.000", CultureInfo.GetCultureInfo("pl-PL"));
                                    //acText.TextString = poziom.ToString();
                                }
                            }
                            acText.TextStyleId = stylwymiarow;
                            acText.HorizontalMode = zzd.TextHorizontalMode.TextLeft;
                            acText.VerticalMode = zzd.TextVerticalMode.TextVerticalMid;
                            if (ptStart.Y < ptEnd.Y)
                            {
                                acText.AlignmentPoint = new zzg.Point3d(ptStart.X - 2 * dimscala, ptStart.Y + dimscala * 5, ptStart.Z);
                            }
                            else
                            {
                                acText.AlignmentPoint = new zzg.Point3d(ptStart.X - 2 * dimscala, ptStart.Y - dimscala * 5, ptStart.Z);
                            }
                            acText.Height = 2 * dimscala;
                            acText.WidthFactor = 0.65;
                            acText.Layer = layername;
                            acText.ColorIndex = 2;
                            acText.TransformBy(mat);

                            btr.AppendEntity(acText);
                            tr.AddNewlyCreatedDBObject(acText, true);
                            anonyGroup.Append(acText.ObjectId);

                        }

                        //rysujemy polilinie trójkąta               
                        using (zzd.Polyline acPoly = new zzd.Polyline())
                        {

                            if (ptStart.Y < ptEnd.Y)
                            {
                                acPoly.AddVertexAt(0, new zzg.Point2d(ptStart.X, ptStart.Y + 7 * dimscala), 0, 0, 0);
                                acPoly.AddVertexAt(1, new zzg.Point2d(ptStart.X, ptStart.Y), 0, 0, 0);
                                acPoly.AddVertexAt(2, new zzg.Point2d(ptStart.X - 2 * dimscala, ptStart.Y + 3 * dimscala), 0, 0, 0);
                                acPoly.AddVertexAt(3, new zzg.Point2d(ptStart.X - 2 * dimscala + tekst_dlugosc, ptStart.Y + 3 * dimscala), 0, 0, 0);

                                acPoly.Layer = layername;
                                acPoly.ColorIndex = 1;
                                acPoly.TransformBy(mat);

                                btr.AppendEntity(acPoly);
                                tr.AddNewlyCreatedDBObject(acPoly, true);
                                anonyGroup.Append(acPoly.ObjectId);
                            }
                            else
                            {
                                acPoly.AddVertexAt(0, new zzg.Point2d(ptStart.X, ptStart.Y - 7 * dimscala), 0, 0, 0);
                                acPoly.AddVertexAt(1, new zzg.Point2d(ptStart.X, ptStart.Y), 0, 0, 0);
                                acPoly.AddVertexAt(2, new zzg.Point2d(ptStart.X - 2 * dimscala, ptStart.Y - 3 * dimscala), 0, 0, 0);
                                acPoly.AddVertexAt(3, new zzg.Point2d(ptStart.X - 2 * dimscala + tekst_dlugosc, ptStart.Y - 3 * dimscala), 0, 0, 0);

                                acPoly.Layer = layername;
                                acPoly.ColorIndex = 1;
                                acPoly.TransformBy(mat);

                                btr.AppendEntity(acPoly);
                                tr.AddNewlyCreatedDBObject(acPoly, true);
                                anonyGroup.Append(acPoly.ObjectId);
                            }
                        }
                        tr.AddNewlyCreatedDBObject(anonyGroup, true);
                        // Commit the changes and dispose of the transaction
                        tr.Commit();
                    }                   
                // Exit if the user presses ESC or cancels the command
                if (pPtRes.Status == zze.PromptStatus.Cancel)
                { 
                    //przełącz zmienną osmode po staremu
                    zza.Application.SetSystemVariable("Osmode", osmode);
                    return;
                }

            }
        }        

        /// <summary>
        /// Skaluje okno widokowe(viewport) zgodnie z aktualną skalą ustawioną programem skala.dll
        /// </summary>
        public static void SkalujViewport()
        {
            

            zza.Document doc = zza.Application.DocumentManager.MdiActiveDocument;
            zzd.Database db = doc.Database;
            zze.Editor ed = doc.Editor;
            zze.PromptSelectionOptions wskazsele = new zze.PromptSelectionOptions();
            zze.PromptSelectionResult selectionResult = ed.SelectImplied();

            //ZAPISZEMY DIMSCALE
            double dimscala = Convert.ToDouble(zza.Application.GetSystemVariable("DIMSCALE"));

            //sprawdzamy czy jesteśmy w przestrzeni papieru
            bool czypapier = ItisModel();
            if (czypapier == true)
            {
                doc.Editor.WriteMessage("\n...Ten moduł działa tylko w przestrzeni papieru");
            }
            else
            {
                wskazsele.MessageForAdding = "\nWskaż rzutnie(viewport): ";
                wskazsele.SingleOnly = true;
                selectionResult = doc.Editor.GetSelection(wskazsele);

                if (selectionResult.Status == zze.PromptStatus.Cancel) return;

                if (selectionResult.Status == zze.PromptStatus.OK)
                {

                    using (zzd.Transaction tr = doc.Database.TransactionManager.StartTransaction())
                    {
                        zzd.BlockTable bt;
                        zzd.BlockTableRecord btr;

                        // Open Model space for write
                        bt = tr.GetObject(db.BlockTableId, zzd.OpenMode.ForRead) as zzd.BlockTable;
                        btr = tr.GetObject(bt[zzd.BlockTableRecord.PaperSpace], zzd.OpenMode.ForWrite) as zzd.BlockTableRecord;

                        zze.SelectionSet currentlySelectedEntities = selectionResult.Value;

                        foreach (zzd.ObjectId id in currentlySelectedEntities.GetObjectIds())
                        {
                            zzd.Entity ent = tr.GetObject(id, zzd.OpenMode.ForWrite) as zzd.Entity;
                            string typ = ent.ToString();
                            //MessageBox.Show(typ);

                            if (typ == "ZwSoft.ZwCAD.DatabaseServices.Viewport")
                            {
                                zzd.Viewport okno = ent as zzd.Viewport;
                                okno.CustomScale = 1.0 / dimscala;
                            }
                            else
                            {
                                doc.Editor.WriteMessage("\n...Wskazany obiekt to nie rzutnia(viewport)");
                            }
                        }
                        tr.Commit();
                    }
                }
                else
                {
                    doc.Editor.WriteMessage("\n...Nie trafiłeś w obiekt");
                }

            }
        }

        /// <summary>
        /// Polilinia typu zigzag
        /// </summary>
        public static void Zigzag()
        {
            for (; ; )
            {
                string layername = "PI_ZIGZAG";
                string startpromt = "\nPodaj punkt startowy : ";
                string endpromt = "\nPodaj punkt końcowy: ";
                //tworzymy odpowiednią warstwę o ile nie istnieje
               
                WstawWarstwe(layername, 1, true);
                //pobieramy bazę danych aktualnego rysunku
                zza.Document doc = zza.Application.DocumentManager.MdiActiveDocument;
                zzd.Database db = doc.Database;
                zze.Editor ed = doc.Editor;

                //sprawdza zmienna tilemode
                int tilemode = Convert.ToInt16(zza.Application.GetSystemVariable("TILEMODE"));
                //zapiszemy aktualną zmienną osmode
                int osmode = Convert.ToInt16(zza.Application.GetSystemVariable("OSMODE"));
                //ZAPISZEMY DIMSCALE
                double dimscala = Convert.ToDouble(zza.Application.GetSystemVariable("DIMSCALE"));
                bool jestmodel = ItisModel();
                bool jestviewport = IsInLayoutPaper();

                if (jestmodel == false & jestviewport == true) dimscala = 1.0;

                zze.PromptPointResult pPtRes;
                zze.PromptPointOptions pPtOpts = new zze.PromptPointOptions("");

                //przełącz zmienną osmode na 512
                zza.Application.SetSystemVariable("Osmode", 512);

                //zapytaj o punkt początkowy
                pPtOpts.Message = startpromt;
                pPtRes = ed.GetPoint(pPtOpts);
                zzg.Point3d ptStart = pPtRes.Value;

                // Exit if the user presses ESC or cancels the command
                if (pPtRes.Status == zze.PromptStatus.Cancel)
                {
                    zza.Application.SetSystemVariable("Osmode", osmode);
                    return;
                }
                //przełącz zmienną osmode na 183
                zza.Application.SetSystemVariable("Osmode", 183);

                // Prompt for the end point
                pPtOpts.Message = endpromt;
                pPtOpts.UseBasePoint = true;
                pPtOpts.BasePoint = ptStart;
                pPtRes = doc.Editor.GetPoint(pPtOpts);
                zzg.Point3d ptEnd = pPtRes.Value;

                if (pPtRes.Status == zze.PromptStatus.Cancel)
                {
                    zza.Application.SetSystemVariable("Osmode", osmode);
                    return;
                }
                zza.Application.SetSystemVariable("Osmode", osmode);

                //conversja
                zzg.Matrix3d ucs = ed.CurrentUserCoordinateSystem;
                zzg.CoordinateSystem3d cs = ucs.CoordinateSystem3d;

                // Transform from UCS to WCS
                zzg.Matrix3d mat = zzg.Matrix3d.AlignCoordinateSystem(zzg.Point3d.Origin, zzg.Vector3d.XAxis, zzg.Vector3d.YAxis, zzg.Vector3d.ZAxis,
                    cs.Origin, cs.Xaxis, cs.Yaxis, cs.Zaxis);

                ptStart = ptStart.TransformBy(mat);
                ptEnd = ptEnd.TransformBy(mat);

                //ustalenie vectora
                int powtorz = 1;
                zzg.Point2d startpoint = new zzg.Point2d(ptStart.X, ptStart.Y);

                zzg.Vector2d vectorpodstawowy = new zzg.Point2d(ptEnd.X, ptEnd.Y) - startpoint;
                zzg.Vector2d vectotjednostkowy = (vectorpodstawowy * 1 / vectorpodstawowy.Length) * dimscala;
                zzg.Vector2d vectotjednostkowy90 = vectotjednostkowy.RotateBy(Math.PI / 2);
                zzg.Point2d startpoint0 = startpoint;
                zzg.Point2d startpoint1 = startpoint0 + 3.5 * vectotjednostkowy;
                zzg.Point2d startpoint2 = startpoint0 + 4.0 * vectotjednostkowy + 1 * vectotjednostkowy90;
                zzg.Point2d startpoint3 = startpoint0 + 4.5 * vectotjednostkowy - 1 * vectotjednostkowy90;

                using (zzd.Transaction tr = db.TransactionManager.StartTransaction())
                {
                    zzd.BlockTable bt;
                    zzd.BlockTableRecord btr;

                    // Open Model space for write
                    bt = tr.GetObject(db.BlockTableId,zzd.OpenMode.ForRead) as zzd.BlockTable;

                    if (jestmodel == true || jestviewport == false)
                    {
                        btr = tr.GetObject(bt[zzd.BlockTableRecord.ModelSpace], zzd.OpenMode.ForWrite) as zzd.BlockTableRecord;
                    }
                    else
                    {
                        btr = tr.GetObject(bt[zzd.BlockTableRecord.PaperSpace], zzd.OpenMode.ForWrite) as zzd.BlockTableRecord;
                    }                    

                    // Create a polyline with two segments (3 points)
                    using (zzd.Polyline acPoly = new zzd.Polyline())
                    {
                        while ((startpoint3 - startpoint).Length < vectorpodstawowy.Length & (startpoint3 - startpoint).Length < vectotjednostkowy.Length * 500.0)
                        {
                            double PRZESUN = 5.0;


                            acPoly.AddVertexAt(powtorz - 1, startpoint0, 0, 0, 0);
                            acPoly.AddVertexAt(powtorz, startpoint1, 0, 0, 0);
                            acPoly.AddVertexAt(powtorz + 1, startpoint2, 0, 0, 0);
                            acPoly.AddVertexAt(powtorz + 2, startpoint3, 0, 0, 0);

                            powtorz = powtorz + 4;

                            startpoint0 = startpoint0 + PRZESUN * vectotjednostkowy;
                            startpoint1 = startpoint1 + PRZESUN * vectotjednostkowy;
                            startpoint2 = startpoint2 + PRZESUN * vectotjednostkowy;
                            startpoint3 = startpoint3 + PRZESUN * vectotjednostkowy;

                        }

                        acPoly.AddVertexAt(powtorz - 1, startpoint0, 0, 0, 0);
                        acPoly.AddVertexAt(powtorz, startpoint0 + 3.5 * vectotjednostkowy, 0, 0, 0);
                        acPoly.Layer = layername;

                        // Add the new object to the block table record and the transaction
                        btr.AppendEntity(acPoly);
                        tr.AddNewlyCreatedDBObject(acPoly, true);
                    }

                    // Commit the changes and dispose of the transaction
                    tr.Commit();

                    if (pPtRes.Status == zze.PromptStatus.Cancel)
                    {                        
                        return;
                    }
                }
            }

        }

        /// <summary>
        /// Polilinia typu izolacja
        /// </summary>
        public static void RysujIzo1()
        {
            for (; ; )
            {
                string layername = "PI_ZIGZAG";
                string startpromt = "\nWskaż początek izolacji: ";
                string midpromt = "\nWskaż punkt wyznaczający grubość: ";
                string endpromt = "\nWskaż zasięg: ";
                //tworzymy odpowiednią warstwę o ile nie istnieje
                WstawWarstwe(layername, 7, true);
                //pobieramy bazę danych aktualnego rysunku
                zza.Document doc = zza.Application.DocumentManager.MdiActiveDocument;
                zzd.Database db = doc.Database;
                zze.Editor ed = doc.Editor;

                //zapiszemy aktualną zmienną osmode
                int osmode = Convert.ToInt16(zza.Application.GetSystemVariable("OSMODE"));
                //ZAPISZEMY DIMSCALE
                double dimscala = Convert.ToDouble(zza.Application.GetSystemVariable("DIMSCALE"));
                bool jestmodel = ItisModel();
                bool jestviewport = IsInLayoutPaper();

                zze.PromptPointResult pPtRes;
                zze.PromptPointOptions pPtOpts = new zze.PromptPointOptions("");

                //przełącz zmienną osmode na 512
                zza.Application.SetSystemVariable("Osmode", 512);

                // zapytaj o srodek
                pPtOpts.Message = startpromt;
                pPtRes = doc.Editor.GetPoint(pPtOpts);
                zzg.Point3d ptStart = pPtRes.Value;

                // Exit if the user presses ESC or cancels the command
                if (pPtRes.Status == zze.PromptStatus.Cancel)
                {
                    zza.Application.SetSystemVariable("Osmode", osmode);
                    return;
                }

                //przełącz zmienną osmode na 128
                zza.Application.SetSystemVariable("Osmode", 128);

                // zapytaj o promień
                pPtOpts.Message = midpromt;
                pPtOpts.UseBasePoint = true;
                pPtOpts.BasePoint = ptStart;
                pPtRes = doc.Editor.GetPoint(pPtOpts);
                zzg.Point3d ptEndgr = pPtRes.Value;

                if (pPtRes.Status == zze.PromptStatus.Cancel)
                {
                    zza.Application.SetSystemVariable("Osmode", osmode);
                    return;
                }

                zzg.Vector3d vectorpodstawowygr = ptEndgr - ptStart;
                double grubosc = vectorpodstawowygr.Length;
                
                // Exit if the user presses ESC or cancels the command
                if (pPtRes.Status == zze.PromptStatus.Cancel) return;

                //przełącz zmienną osmode na 512
                zza.Application.SetSystemVariable("Osmode", 512);

                // zapytaj o promień
                pPtOpts.Message = endpromt;
                pPtOpts.UseBasePoint = true;
                pPtOpts.BasePoint = ptStart;
                pPtRes = doc.Editor.GetPoint(pPtOpts);
                zzg.Point3d ptEnd = pPtRes.Value;

                if (pPtRes.Status == zze.PromptStatus.Cancel)
                {
                    zza.Application.SetSystemVariable("Osmode", osmode);
                    return;
                }

                //przełącz zmienną osmode na ZAPISANA
                zza.Application.SetSystemVariable("Osmode", osmode);

                //conversja
                zzg.Matrix3d ucs = ed.CurrentUserCoordinateSystem;
                zzg.CoordinateSystem3d cs = ucs.CoordinateSystem3d;

                // Transform from UCS to WCS
                zzg.Matrix3d mat = zzg.Matrix3d.AlignCoordinateSystem(zzg.Point3d.Origin, zzg.Vector3d.XAxis, zzg.Vector3d.YAxis, zzg.Vector3d.ZAxis,
                    cs.Origin, cs.Xaxis, cs.Yaxis, cs.Zaxis);

                ptStart = ptStart.TransformBy(mat);
                ptEnd = ptEnd.TransformBy(mat);

                //ustalenie vectora   
                int powtorz = 1;
                zzg.Point2d startpoint = new zzg.Point2d(ptStart.X, ptStart.Y);

                zzg.Vector2d vectorpodstawowy = new zzg.Point2d(ptEnd.X, ptEnd.Y) - startpoint;
                zzg.Vector2d vectotjednostkowy = (vectorpodstawowy * 1 / vectorpodstawowy.Length) * grubosc;
                zzg.Vector2d vectotjednostkowy90 = vectotjednostkowy.RotateBy(Math.PI / 2);
                zzg.Point2d startpoint0 = startpoint;
                zzg.Point2d startpoint1 = startpoint0 + 0.25 * vectotjednostkowy + 0.25 * vectotjednostkowy90;
                zzg.Point2d startpoint2 = startpoint0 + 0.75 * vectotjednostkowy90;
                zzg.Point2d startpoint3 = startpoint0 + 0.25 * vectotjednostkowy + vectotjednostkowy90;
                zzg.Point2d startpoint4 = startpoint0 + 0.5 * vectotjednostkowy + 0.75 * vectotjednostkowy90;
                zzg.Point2d startpoint5 = startpoint1;

                using (zzd.Transaction tr = db.TransactionManager.StartTransaction())
                {
                    zzd.BlockTable bt;
                    zzd.BlockTableRecord btr;

                    // Open Model space for write
                    bt = tr.GetObject(db.BlockTableId,
                                                    zzd.OpenMode.ForRead) as zzd.BlockTable;

                    if (jestmodel == true || jestviewport == false)
                    {
                        btr = tr.GetObject(bt[zzd.BlockTableRecord.ModelSpace], zzd.OpenMode.ForWrite) as zzd.BlockTableRecord;
                    }
                    else
                    {
                        btr = tr.GetObject(bt[zzd.BlockTableRecord.PaperSpace], zzd.OpenMode.ForWrite) as zzd.BlockTableRecord;
                    }                   

                    using (zzd.Polyline acPoly = new zzd.Polyline())
                    {
                        double PRZESUN = 0.5;
                        while ((startpoint5 - startpoint).Length < grubosc * 200.0 & (startpoint5 - startpoint).Length < vectorpodstawowy.Length)
                        {


                            acPoly.AddVertexAt(powtorz - 1, startpoint0, 0.4206, 0, 0);
                            acPoly.AddVertexAt(powtorz, startpoint1, 0, 0, 0);
                            acPoly.AddVertexAt(powtorz + 1, startpoint2, -0.4206, 0, 0);
                            acPoly.AddVertexAt(powtorz + 2, startpoint3, -0.4206, 0, 0);
                            acPoly.AddVertexAt(powtorz + 3, startpoint4, 0, 0, 0);
                            acPoly.AddVertexAt(powtorz + 4, startpoint5, 0.4206, 0, 0);

                            powtorz = powtorz + 6;

                            startpoint0 = startpoint0 + PRZESUN * vectotjednostkowy;
                            startpoint1 = startpoint1 + PRZESUN * vectotjednostkowy;
                            startpoint2 = startpoint2 + PRZESUN * vectotjednostkowy;
                            startpoint3 = startpoint3 + PRZESUN * vectotjednostkowy;
                            startpoint4 = startpoint4 + PRZESUN * vectotjednostkowy;
                            startpoint5 = startpoint5 + PRZESUN * vectotjednostkowy;
                        }

                        acPoly.AddVertexAt(powtorz - 1, startpoint0, 0, 0, 0);
                        acPoly.Layer = layername;


                        // Add the new object to the block table record and the transaction
                        btr.AppendEntity(acPoly);
                        tr.AddNewlyCreatedDBObject(acPoly, true);
                    }
                    zza.Application.SetSystemVariable("Osmode", osmode);
                    // Commit the changes and dispose of the transaction
                    tr.Commit();

                    if (pPtRes.Status == zze.PromptStatus.Cancel)
                    {                        
                        return;
                    }
                }
            }
        }
    }
}
