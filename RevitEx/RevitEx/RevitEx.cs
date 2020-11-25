using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using RevitDB = Autodesk.Revit.DB;
using RevitUI = Autodesk.Revit.UI;
using RevitEvents = Autodesk.Revit.UI.Events;
using RevitAttributes = Autodesk.Revit.Attributes;
using RevitApplicationServices = Autodesk.Revit.ApplicationServices;




namespace RevitEx
{
    [RevitAttributes.Transaction(RevitAttributes.TransactionMode.Manual)]
    [RevitAttributes.Regeneration(RevitAttributes.RegenerationOption.Manual)]
    public class App : RevitUI.IExternalApplication
    {
        static RevitDB.AddInId appId = new RevitDB.AddInId(new Guid("369A57E5-9113-4676-AA0D-F358914569C1"));
        // get the absolute path of this assembly
        static string appAssembly = System.Reflection.Assembly.GetExecutingAssembly().Location;
        static string appAssemblyPath = appAssembly.Substring(0, appAssembly.LastIndexOf(@"\"));
        private AppDocEvents appDocEvents;

        public RevitUI.Result OnStartup(RevitUI.UIControlledApplication application)
        {
            AddRibbon(application);
            AddAppDocEvents(application.ControlledApplication);
            return RevitUI.Result.Succeeded;
        }

        public RevitUI.Result OnShutdown(RevitUI.UIControlledApplication application)
        {
            RemoveAppDocEvents();

            return RevitUI.Result.Succeeded;
        }

        private void AddRibbon(RevitUI.UIControlledApplication app)
        {
            app.CreateRibbonTab("数据接口");

            RevitUI.RibbonPanel ribbon_panel = app.CreateRibbonPanel("数据接口", "数据");
            RevitUI.PulldownButtonData data_pull = new RevitUI.PulldownButtonData("RevitMethod", "功能");
            RevitUI.PulldownButton btn_pull = ribbon_panel.AddItem(data_pull) as RevitUI.PulldownButton;
            btn_pull.LargeImage = new System.Windows.Media.Imaging.BitmapImage(new Uri(appAssemblyPath + @"\Revit\RevitEx.png"));
            btn_pull.AddPushButton(new RevitUI.PushButtonData("TestDlg", "Hello World", appAssembly, "RevitEx.cmdTest"));
            btn_pull.AddPushButton(new RevitUI.PushButtonData("Journaling", "Objects Journaling", appAssembly, "RevitEx.cmdJournaling"));
            btn_pull.AddPushButton(new RevitUI.PushButtonData("ShowObjects", "Objects Show", appAssembly, "RevitEx.cmdShowSteels"));

            ribbon_panel = app.CreateRibbonPanel("数据接口", "接口");
            RevitUI.SplitButtonData data_split = new RevitUI.SplitButtonData("RevitExcel", "Excel接口");
            RevitUI.SplitButton btn_split = ribbon_panel.AddItem(data_split) as RevitUI.SplitButton;
            btn_split.LargeImage = new System.Windows.Media.Imaging.BitmapImage(new Uri(appAssemblyPath + @"\Revit\RevitExcel.png"));
            RevitUI.PushButton btn_push = btn_split.AddPushButton(new RevitUI.PushButtonData("ExportExcel", "导出Excel", appAssembly, "RevitEx.cmdExportExcel"));
            btn_push.LargeImage = new System.Windows.Media.Imaging.BitmapImage(new Uri(appAssemblyPath + @"\Revit\ExportExcel.png"));
            btn_push = btn_split.AddPushButton(new RevitUI.PushButtonData("ImportExcel", "导入Excel", appAssembly, "RevitEx.cmdImportExcel"));
            btn_push.LargeImage = new System.Windows.Media.Imaging.BitmapImage(new Uri(appAssemblyPath + @"\Revit\ImportExcel.png"));

            //创建下拉组合框
            ribbon_panel = app.CreateRibbonPanel("数据接口", "控件");
            RevitUI.ComboBoxData data_combo = new RevitUI.ComboBoxData("选项");
            RevitUI.ComboBox cbx = ribbon_panel.AddItem(data_combo) as RevitUI.ComboBox;

            if (cbx != null)
            {
                cbx.ItemText = "选择操作";

                RevitUI.ComboBoxMemberData data_cbxm = new RevitUI.ComboBoxMemberData("Close", "关闭");
                data_cbxm.GroupName = "编辑操作";
                cbx.AddItem(data_cbxm);
                data_cbxm = new RevitUI.ComboBoxMemberData("Change", "修改");
                cbx.AddItem(data_cbxm);
            }
            cbx.CurrentChanged += change;
            cbx.DropDownClosed += closed;
        }

        private void AddMenu(RevitUI.UIControlledApplication app)
        {
            RevitUI.RibbonPanel ribbon_panel = app.CreateRibbonPanel("数据接口");
            RevitUI.PulldownButtonData data_pull = new RevitUI.PulldownButtonData("RevitTest", "测试功能");

            RevitUI.PulldownButton btn_pull = ribbon_panel.AddItem(data_pull) as RevitUI.PulldownButton;
            btn_pull.LargeImage = new System.Windows.Media.Imaging.BitmapImage(new Uri(appAssemblyPath + @"\Revit\RevitEx.png"));

            btn_pull.AddPushButton(new RevitUI.PushButtonData("Test", "Hello World", appAssembly, "RevitEx.cmdTest"));
            btn_pull.AddPushButton(new RevitUI.PushButtonData("Journaling", "Objects Journaling.", appAssembly, "RevitEx.cmdJournaling"));
            btn_pull.AddPushButton(new RevitUI.PushButtonData("ShowObjects", "Objects Show", appAssembly, "RevitEx.cmdShowSteels"));

            RevitUI.PushButtonData data_push = new RevitUI.PushButtonData("RevitExcel", "导出Excel", appAssembly, "RevitEx.cmdExportExcel");
            RevitUI.PushButton btn_push = ribbon_panel.AddItem(data_push) as RevitUI.PushButton;
            btn_push.LargeImage = new System.Windows.Media.Imaging.BitmapImage(new Uri(appAssemblyPath + @"\Revit\RevitExcel.png"));
        }

        private void AddAppDocEvents(RevitApplicationServices.ControlledApplication app)
        {
            appDocEvents = new AppDocEvents(app);
            appDocEvents.EnableEvents();
        }

        private void RemoveAppDocEvents()
        {
            appDocEvents.DisableEvents();
        }

        private void closed(object sender, RevitEvents.ComboBoxDropDownClosedEventArgs e)
        {
            MessageBox.Show("关闭", "已关闭");
        }

        private void change(object sender, RevitEvents.ComboBoxCurrentChangedEventArgs e)
        {
            MessageBox.Show("修改", "已修改");
        }
    }

    /// <summary>
    /// 定义Transaction特性。这个特性源自Autodesk.Revit.Attributes.TransactionAttribute。该特性有三种模式：Automatic（自动）、Manual（手动）和ReadOnly（只读）。因为该特性没有默认值，所以必须显式指定。在本例中，模式可任选。
    /// 如果觉得这句太长，可以在“using……”代码块加上using Autodesk.Revit.Attributes;
    /// 这句就可以写成[Transaction(TransactionMode.Manual)]
    /// 声明一个类，继承RevitAPI的IExternalCommand（外部命令）接口。
    /// 重载Execute()函数。可以把它粗略的理解为IExternalCommand接口类的主函数或入口函数，类似于Java里的main()函数那样的东西。这个函数被Autodesk.Revit.UI.Result限制，所以必须有返回值。
    /// 返回Succeeded。Autodesk.Revit.UI.Result有三个值，分别是Succeeded、Failed和Canceled。如果没有返回Succeeded，Revit会撤销所做的操作。
    /// </summary>
    [RevitAttributes.Transaction(RevitAttributes.TransactionMode.Manual)]
    public class cmdTest : RevitUI.IExternalCommand
    {
        public RevitUI.Result Execute(RevitUI.ExternalCommandData commandData, ref string message, RevitDB.ElementSet elements)
        {
            RevitUI.TaskDialog.Show("Revit", "Hello Test World!");
            return RevitUI.Result.Succeeded;
        }
    }

    [RevitAttributes.Transaction(RevitAttributes.TransactionMode.Manual)]
    public class cmdExportExcel : RevitUI.IExternalCommand
    {
        public RevitUI.Result Execute(RevitUI.ExternalCommandData commandData, ref string message, RevitDB.ElementSet elements)
        {
            RevitUI.UIDocument uidoc = commandData.Application.ActiveUIDocument;
            RevitDB.Document doc = uidoc.Document;
            if (doc.IsFamilyDocument)//感觉族文档没必要
            {
                RevitUI.TaskDialog.Show("Revit", "该操作仅适用项目文档，不适用族文档！");
                return RevitUI.Result.Succeeded;
            }
            try
            {
                RevitDB.ViewSchedule vs = doc.ActiveView as RevitDB.ViewSchedule;
                RevitDB.TableData td = vs.GetTableData();
                RevitDB.TableSectionData tsdh = td.GetSectionData(RevitDB.SectionType.Header);
                RevitDB.TableSectionData tsdb = td.GetSectionData(RevitDB.SectionType.Body);

                Excel.Application app = new Excel.Application();
                Excel.Workbooks wbs = app.Workbooks;
                Excel.Workbook wb = wbs.Add(Type.Missing);
                Excel.Worksheet ws = wb.Worksheets["Sheet1"];

                int cs = tsdb.NumberOfColumns;
                int rs = tsdb.NumberOfRows;

                Excel.Range rg1 = ws.Cells[1, 1];
                Excel.Range rg2 = ws.Cells[1, cs];
                rg1.Value = vs.GetCellText(RevitDB.SectionType.Header, 0, 0);
                rg2.Value = "";
                Excel.Range rg = ws.get_Range(rg1, rg2);
                rg.Merge();
                rg.Font.Name = "黑体";
                rg.Font.Bold = 400;
                rg.Font.Size = 14;
                rg.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                rg.HorizontalAlignment = Excel.Constants.xlCenter;
                rg.RowHeight = 25;

                for (int i = 0; i < rs; i++)
                {
                    for (int j = 0; j < cs; j++)
                    {
                        RevitDB.CellType ct = tsdb.GetCellType(i, j);
                        ws.Cells[i + 2, j + 1] = vs.GetCellText(RevitDB.SectionType.Body, i, j);
                    }
                }

                rg1 = ws.Cells[2, 1];
                rg2 = ws.Cells[rs + 1, cs];
                rg = ws.get_Range(rg1, rg2);
                rg.Font.Name = "仿宋";
                rg.Font.Size = 11;
                rg.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                rg.HorizontalAlignment = Excel.Constants.xlCenter;
                rg.Borders.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                rg.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                rg.EntireColumn.AutoFit();
                rg.RowHeight = 20;

                rg.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium;
                rg.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).ColorIndex = 3;
                rg.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium;
                rg.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).ColorIndex = 3;
                rg.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium;
                rg.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).ColorIndex = 3;
                rg.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium;
                rg.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).ColorIndex = 3;

                SaveFileDialog sf = new SaveFileDialog();
                sf.Title = "Revit";
                sf.AddExtension = true;
                sf.Filter = "Excel2007文件|*.xlsx|Excel2003文件|*.xls|文本文件|*.txt|所有文件|*.*";
                sf.FileName = vs.ViewName;
                if (DialogResult.OK == sf.ShowDialog())
                    wb.SaveAs(sf.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                app.AlertBeforeOverwriting = false;
                wb.Close(true, Type.Missing, Type.Missing);
                wbs.Close();
                app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                wb = null;
                wbs = null;
                app = null;
                GC.Collect();
            }
            catch (Exception)
            {
                return RevitUI.Result.Cancelled;
            }

            RevitUI.TaskDialog.Show("Revit", "Revit Export Excel Successful!");
            return RevitUI.Result.Succeeded;
        }
    }

    [RevitAttributes.Transaction(RevitAttributes.TransactionMode.Manual)]
    public class cmdImportExcel : RevitUI.IExternalCommand
    {
        public RevitUI.Result Execute(RevitUI.ExternalCommandData commandData, ref string message, RevitDB.ElementSet elements)
        {
            RevitUI.UIDocument uidoc = commandData.Application.ActiveUIDocument;
            RevitDB.Document doc = uidoc.Document;
            if (doc.IsFamilyDocument)//感觉族文档没必要
            {
                RevitUI.TaskDialog.Show("Revit", "该操作仅适用项目文档，不适用族文档！");
                return RevitUI.Result.Succeeded;
            }
            try
            {
                RevitDB.ViewSchedule vs = doc.ActiveView as RevitDB.ViewSchedule;
                RevitDB.TableData td = vs.GetTableData();
                RevitDB.TableSectionData tsdh = td.GetSectionData(RevitDB.SectionType.Header);
                RevitDB.TableSectionData tsdb = td.GetSectionData(RevitDB.SectionType.Body);

                Excel.Application app = new Excel.Application();
                Excel.Workbooks wbs = app.Workbooks;
                Excel.Workbook wb = wbs.Add(Type.Missing);
                Excel.Worksheet ws = wb.Worksheets["Sheet1"];

                int cs = tsdb.NumberOfColumns;
                int rs = tsdb.NumberOfRows;

                Excel.Range rg1 = ws.Cells[1, 1];
                Excel.Range rg2 = ws.Cells[1, cs];
                rg1.Value = vs.GetCellText(RevitDB.SectionType.Header, 0, 0);
                rg2.Value = "";
                Excel.Range rg = ws.get_Range(rg1, rg2);
                rg.Merge();
                rg.Font.Name = "黑体";
                rg.Font.Bold = 400;
                rg.Font.Size = 14;
                rg.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                rg.HorizontalAlignment = Excel.Constants.xlCenter;
                rg.RowHeight = 25;

                for (int i = 0; i < rs; i++)
                {
                    for (int j = 0; j < cs; j++)
                    {
                        RevitDB.CellType ct = tsdb.GetCellType(i, j);
                        ws.Cells[i + 2, j + 1] = vs.GetCellText(RevitDB.SectionType.Body, i, j);
                    }
                }

                rg1 = ws.Cells[2, 1];
                rg2 = ws.Cells[rs + 1, cs];
                rg = ws.get_Range(rg1, rg2);
                rg.Font.Name = "仿宋";
                rg.Font.Size = 11;
                rg.Font.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                rg.HorizontalAlignment = Excel.Constants.xlCenter;
                rg.Borders.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
                rg.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                rg.EntireColumn.AutoFit();
                rg.RowHeight = 20;

                rg.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).Weight = Excel.XlBorderWeight.xlMedium;
                rg.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).ColorIndex = 3;
                rg.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).Weight = Excel.XlBorderWeight.xlMedium;
                rg.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).ColorIndex = 3;
                rg.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).Weight = Excel.XlBorderWeight.xlMedium;
                rg.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).ColorIndex = 3;
                rg.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).Weight = Excel.XlBorderWeight.xlMedium;
                rg.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).ColorIndex = 3;

                SaveFileDialog sf = new SaveFileDialog();
                sf.Title = "Revit";
                sf.AddExtension = true;
                sf.Filter = "Excel2007文件|*.xlsx|Excel2003文件|*.xls|文本文件|*.txt|所有文件|*.*";
                sf.FileName = vs.ViewName;
                if (DialogResult.OK == sf.ShowDialog())
                    wb.SaveAs(sf.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                app.AlertBeforeOverwriting = false;
                wb.Close(true, Type.Missing, Type.Missing);
                wbs.Close();
                app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                wb = null;
                wbs = null;
                app = null;
                GC.Collect();
            }
            catch (Exception)
            {
                return RevitUI.Result.Cancelled;
            }

            RevitUI.TaskDialog.Show("Revit", "Revit Export Excel Successful!");
            return RevitUI.Result.Succeeded;
        }
    }

    [RevitAttributes.Transaction(RevitAttributes.TransactionMode.Manual)]
    [RevitAttributes.Regeneration(RevitAttributes.RegenerationOption.Manual)]
    [RevitAttributes.Journaling(RevitAttributes.JournalingMode.NoCommandData)]
    public class cmdJournaling : RevitUI.IExternalCommand
    {
        public RevitUI.Result Execute(RevitUI.ExternalCommandData commandData, ref string message, RevitDB.ElementSet elements)
        {
            try
            {
                RevitDB.Transaction tran = new RevitDB.Transaction(commandData.Application.ActiveUIDocument.Document, "Journaling");
                tran.Start();
                Journaling deal = new Journaling(commandData);
                deal.Run();
                tran.Commit();

                return RevitUI.Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return RevitUI.Result.Failed;
            }
        }
    }

    [RevitAttributes.Transaction(RevitAttributes.TransactionMode.Manual)]
    [RevitAttributes.Regeneration(RevitAttributes.RegenerationOption.Manual)]
    [RevitAttributes.Journaling(RevitAttributes.JournalingMode.NoCommandData)]
    public class cmdShowObjects : RevitUI.IExternalCommand //外部命令
    {
        public RevitUI.Result Execute(RevitUI.ExternalCommandData commandData, ref string message, RevitDB.ElementSet elements)
        {
            RevitUI.UIDocument uidoc = commandData.Application.ActiveUIDocument;
            RevitDB.Document doc = uidoc.Document;
            if (doc.IsFamilyDocument)//感觉族文档没必要
            {
                RevitUI.TaskDialog.Show("Revit", "该操作仅适用项目文档，不适用族文档！");
                return RevitUI.Result.Succeeded;
            }
            try
            {
                RevitDB.View view = doc.ActiveView;//当前视图
                ICollection<RevitDB.ElementId> ids = uidoc.Selection.GetElementIds();//选择对象
                if (0 == ids.Count)//如果没有选择
                {
                    if (view is RevitDB.ViewSection)//当前视图为立面视图且没有选择对象时退出
                    {
                        return RevitUI.Result.Succeeded;
                    }
                    else if (view is RevitDB.View3D)//当前视图为三维视图且没有选择对象时切换至原楼层平面视图
                    {
                        RevitDB.FilteredElementCollector collectL = new RevitDB.FilteredElementCollector(doc);
                        ICollection<RevitDB.Element> collectionL = collectL.OfClass(typeof(RevitDB.View)).ToElements();
                        foreach (RevitDB.Element element in collectionL)
                        {
                            RevitDB.View v = element as RevitDB.View;
                            if (view.ViewName == "三维视图_" + v.ViewType.ToString() + "_" + v.ViewName)
                            {
                                if (!v.IsTemplate)
                                {
                                    uidoc.ActiveView = v;
                                    break;
                                }
                            }
                        }
                        return RevitUI.Result.Succeeded;
                    }//未选择对象时列出当前视图所有可见对象
                    RevitDB.FilteredElementCollector collector = new RevitDB.FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType();
                    ids = collector.ToElementIds();
                }
                RevitDB.Transaction ts = new RevitDB.Transaction(doc, "Create_View3D");
                ts.Start();
                bool fg = true;
                RevitDB.FilteredElementCollector collect = new RevitDB.FilteredElementCollector(doc);
                ICollection<RevitDB.Element> collection = collect.OfClass(typeof(RevitDB.View3D)).ToElements();
                RevitDB.View3D view3d = null;
                foreach (RevitDB.Element element in collection)
                {
                    RevitDB.View3D v = element as RevitDB.View3D;
                    if (v.ViewName == "三维视图_" + view.ViewType.ToString() + "_" + view.ViewName)
                    {
                        if (!v.IsTemplate)
                        {
                            view3d = v;
                            fg = false;//已生成过当前视图的三维视图
                            break;
                        }
                    }
                }
                if (fg)
                {
                    RevitDB.ViewFamilyType viewFamilyType = null;
                    collect = new RevitDB.FilteredElementCollector(doc);
                    var viewFamilyTypes = collect.OfClass(typeof(RevitDB.ViewFamilyType)).ToElements();
                    foreach (RevitDB.Element e in viewFamilyTypes)
                    {
                        RevitDB.ViewFamilyType v = e as RevitDB.ViewFamilyType;
                        if (v.ViewFamily == RevitDB.ViewFamily.ThreeDimensional)
                        {
                            viewFamilyType = v;
                            break;
                        }
                    }
                    view3d = RevitDB.View3D.CreateIsometric(doc, viewFamilyType.Id);
                    view3d.ViewName = "三维视图_" + view.ViewType.ToString() + "_" + view.ViewName;
                }
                ts.Commit();
                uidoc.ActiveView = view3d;//设置生成的三维视图为当前视图
                RevitDB.Transaction tss = new RevitDB.Transaction(doc, "View3D_HideElements");
                tss.Start();
                if (fg)
                {
                    view = doc.ActiveView;
                    collect = new RevitDB.FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType();
                    ICollection<RevitDB.ElementId> idss = collect.ToElementIds();
                    ICollection<RevitDB.ElementId> ds = idss;
                    ds = idss.Except(ids).ToList();
                    ids.Clear();
                    foreach (RevitDB.ElementId id in ds)
                    {
                        RevitDB.Element element = doc.GetElement(id);
                        if (element.CanBeHidden(view) == true)
                        {
                            ids.Add(id);
                        }
                    }
                    view.HideElementsTemporary(ids);//临时隐藏其他图元
                }
                tss.Commit();
            }
            catch (Exception)
            {
                return RevitUI.Result.Cancelled;
            }
            return RevitUI.Result.Succeeded;
        }
    }
}