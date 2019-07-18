using System;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;
using AutoCAD;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Colors;

namespace Wxx_CAD
{

    public static  class WxxTU
    {
        /// <summary>
        /// 用来操作当前视图的函数  
        /// <param name="pMin"></param>
        /// <param name="pMax"></param>
        /// <param name="pCenter"></param>
        /// <param name="dFactor"></param>
        public static void Zoom(Point3d pMin, Point3d pMax, Point3d pCenter, double dFactor)
        {//获取当前文档及数据库
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            int nCurVport = System.Convert.ToInt32(Application.GetSystemVariable("CVPORT"));
            //没提供点或只提供了一个中心点时，获取当前空间的范围
            //检查当前空间是否为模型空间
            if (acCurDb.TileMode == true)
            {
                if (pMin.Equals(new Point3d()) == true &&
                pMax.Equals(new Point3d()) == true)
                {
                    pMin = acCurDb.Extmin;
                    pMax = acCurDb.Extmax;
                }
            }
            else
            {
                //检查当前空间是否为图纸空间
                if (nCurVport == 1)
                {
                    //获取图纸空间范围
                    if (pMin.Equals(new Point3d()) == true &&
                    pMax.Equals(new Point3d()) == true)
                    {
                        pMin = acCurDb.Pextmin;
                        pMax = acCurDb.Pextmax;
                    }
                }
                else
                {
                    //获取模型空间范围
                    if (pMin.Equals(new Point3d()) == true &&
                    pMax.Equals(new Point3d()) == true)
                    {
                        pMin = acCurDb.Extmin;
                        pMax = acCurDb.Extmax;
                    }
                }
            }
            //启动事务
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                //获取当前视图
                using (ViewTableRecord acView = acDoc.Editor.GetCurrentView())
                {
                    Extents3d eExtents;
                    //将WCS坐标变换为DCS坐标
                    Matrix3d matWCS2DCS;
                    matWCS2DCS = Matrix3d.PlaneToWorld(acView.ViewDirection);
                    matWCS2DCS = Matrix3d.Displacement(acView.Target - Point3d.Origin) * matWCS2DCS;
                    matWCS2DCS = Matrix3d.Rotation(-acView.ViewTwist,
                    acView.ViewDirection,
                   acView.Target) * matWCS2DCS;
                    //如果指定了中心点，就为中心模式和比例模式
                    //设置显示范围的最小点和最大点；
                    if (pCenter.DistanceTo(Point3d.Origin) != 0)
                    {
                        pMin = new Point3d(pCenter.X - (acView.Width / 2),
                            pCenter.Y - (acView.Height / 2), 0);
                        pMax = new Point3d((acView.Width / 2) + pCenter.X,
                        (acView.Height / 2) + pCenter.Y, 0);
                    }
                    //用直线创建范围对象；
                    using (Line acLine = new Line(pMin, pMax))
                    {
                        eExtents = new Extents3d(acLine.Bounds.Value.MinPoint,
                        acLine.Bounds.Value.MaxPoint);
                    }
                    //计算当前视图的宽高比
                    double dViewRatio;
                    dViewRatio = (acView.Width / acView.Height);
                    //变换视图范围
                    matWCS2DCS = matWCS2DCS.Inverse();
                    eExtents.TransformBy(matWCS2DCS);
                    double dWidth;
                    double dHeight;
                    Point2d pNewCentPt;
                    //检查是否提供了中心点(中心模式和比例模式)
                    if (pCenter.DistanceTo(Point3d.Origin) != 0)
                    {
                        dWidth = acView.Width;
                        dHeight = acView.Height;
                        if (dFactor == 0)
                        {
                            pCenter = pCenter.TransformBy(matWCS2DCS);
                        }
                        pNewCentPt = new Point2d(pCenter.X, pCenter.Y);
                    }
                    else //窗口、范围和界限模式下
                    {
                        //计算当前视图的宽高新值；
                        dWidth = eExtents.MaxPoint.X - eExtents.MinPoint.X;
                        dHeight = eExtents.MaxPoint.Y - eExtents.MinPoint.Y;
                        //获取视图中心点
                        pNewCentPt = new Point2d(((eExtents.MaxPoint.X + eExtents.MinPoint.X) * 0.5),
                        ((eExtents.MaxPoint.Y + eExtents.MinPoint.Y) * 0.5));
                    }
                    //检查宽度新值是否适于当前窗口
                    if (dWidth > (dHeight * dViewRatio)) dHeight = dWidth / dViewRatio;
                    //调整视图大小；
                    if (dFactor != 0)
                    {
                        acView.Height = dHeight * dFactor;
                        acView.Width = dWidth * dFactor;
                    }
                    //设置视图中心；
                    acView.CenterPoint = pNewCentPt;
                    //更新当前视图；
                    acDoc.Editor.SetCurrentView(acView);
                }
                //提交更改；
                acTrans.Commit();
            }
        }
        /// <summary>
        /// 引线和注释
        /// </summary>
       
        public static void AddLeaderAnnotation(double YB_x,double YB_y,double Z_x,double Z_y,string zhushi
            ,int Color,double G_x,double G_y)
        {
            //获取当前数据库
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            //启动事务
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                //以读模式打开块表
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                OpenMode.ForRead) as BlockTable;
                //以写模式打开块表记录模型空间
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                OpenMode.ForWrite) as BlockTableRecord;
                //创建多行文字(MText)注释
                using (MText acMText = new MText())
                {
                    acMText.Contents = zhushi;
                    acMText.Location = new Point3d(Z_x, Z_y, 0);
                  
                    acMText.ColorIndex = Color;
                    acMText.Width = 2;
                    acMText.Height = 1.5;
                    //添加新对象到模型空间，记录事务
                    acBlkTblRec.AppendEntity(acMText);
                    acTrans.AddNewlyCreatedDBObject(acMText, true);
                    //创建带注释的引线
                    using (Leader acLdr = new Leader())
                    {
                        acLdr.AppendVertex(new Point3d(YB_x,YB_y, 0));
                        acLdr.AppendVertex(new Point3d(G_x, G_y, 0));
                        acLdr.AppendVertex(new Point3d(G_x, G_y+1, 0));
                        acLdr.HasArrowHead = true;
                        //添加新对象到模型空间，记录事务
                        acBlkTblRec.AppendEntity(acLdr);
                        acTrans.AddNewlyCreatedDBObject(acLdr, true);
                        //给引线对象附加注释
                        acLdr.Annotation = acMText.ObjectId;
                        acLdr.EvaluateLeader();
                        //释放DBObject对象
                    }
                }
                //提交修改，回收内存
                acTrans.Commit();
            }
        }

        /// <summary>
        /// 添加线
        /// </summary>
        /// <param name="Chuliqujian"></param>
        /// <param name="B_Sta"></param>
        /// <param name="line_width"></param>
        public static void AddLine(double Chuliqujian, double[] B_Sta, double[] line_width, Document acDoc, Database acCurDb)
        {
            Application.SetSystemVariable("CECOLOR"," 0");
           
            //获取当前文档及数据库
            //  Document acDoc = Application.DocumentManager.MdiActiveDocument;
            //  Database acCurDb = acDoc.Database;
            //绘制横线
            #region
            for (int j = 0; j < line_width.Length; j++)
            {

                //启动事务
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    //以读模式打开块表
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                    OpenMode.ForRead) as BlockTable;
                    //以写模式打开 Block 表记录 Model 空间
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                    OpenMode.ForWrite) as BlockTableRecord;
                    //Create a polyline with two segments (3 points)
                    using (Autodesk.AutoCAD.DatabaseServices.Polyline acPoly = new Autodesk.AutoCAD.DatabaseServices.Polyline())

                    {
                        for (int i = 0; i < B_Sta.Length; i++)

                        {
                            acPoly.AddVertexAt(i, new Point2d(B_Sta[i], line_width[j]), 0, 0, 0);

                        }
                        //将新对象添加到块表记录和事务
                        acBlkTblRec.AppendEntity(acPoly);
                        acTrans.AddNewlyCreatedDBObject(acPoly, true);

                    }
                    //将新对象保存到数据库
                    acTrans.Commit();
                }
            }
            #endregion

            //绘制竖线
            #region
            double Fen = B_Sta[B_Sta.Length - 1] / Chuliqujian;
            for (int i = 0; i < Fen + 1; i++)
            {
                //启动事务
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    //以读模式打开块表
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                    OpenMode.ForRead) as BlockTable;
                    //以写模式打开 Block 表记录 Model 空间
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                    OpenMode.ForWrite) as BlockTableRecord;
                    //以 5,5、12,3两点为起终点创建一条直线
                    using (Line acLine = new Line(new Point3d(B_Sta[0] + i * Chuliqujian, 0, 0),
                    new Point3d(B_Sta[0] + i * Chuliqujian, line_width[line_width.Length - 1], 0)))
                    {
                        //将新对象添加到块表记录和事务
                        acBlkTblRec.AppendEntity(acLine);
                        acTrans.AddNewlyCreatedDBObject(acLine, true);
                        //释放DBObject对象
                    }
                    //将新对象保存到数据库
                    acTrans.Commit();
                }

            }
            #endregion
        }
      
        /// <summary>
        /// 保存文件
        /// </summary>
        public static void SaveActiveDrawing(Document acDoc)
        {

            object obj = Application.GetSystemVariable("DBMOD");
            //检查系统变量DBMOD的值，0表示没有未保存修改
            if (System.Convert.ToInt16(obj) != 0)
            {
                if (System.Windows.Forms.MessageBox.Show("Do you wish to save this drawing?",
                "Save Drawing",
               System.Windows.Forms.MessageBoxButtons.YesNo,
               System.Windows.Forms.MessageBoxIcon.Question)
               == System.Windows.Forms.DialogResult.Yes)
                {
                    string strDWGName=acDoc.Name;
                    //图形命名了吗？0-没呢
                    object obj1 = Application.GetSystemVariable("DWGTITLED");
                    if (System.Convert.ToInt16(obj1) == 0)
                    {
                        //如果图形使用了默认名 (Drawing1、Drawing2等)，
                        //就提供一个新文件名
                        strDWGName = "d:\\MyDrawing.dwg";
                    }
                        acDoc = Application.DocumentManager.MdiActiveDocument;
                        acDoc.Database.SaveAs(strDWGName, true, DwgVersion.Current,
                        acDoc.Database.SecurityParameters);
                    
                }
            }
        }
        /// <summary>
        /// 添加文字
        /// </summary>
        /// <param name="Text"></param>
        /// <param name="hudu"></param>
        /// <param name="colorindex"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="size"></param>
        /// <param name="acDoc"></param>
        /// <param name="acCurDb"></param>
        public static void ObliqueText(string Text, double hudu, int colorindex, double x, double y, double size, Document acDoc, Database acCurDb)
        {
            //double[] B_Sta = new double[] { 0, 2, 10, 30, 40, 50, 60 };
            //   double Chuliqujian = 5;
            // double Fen = B_Sta[B_Sta.Length - 1] / Chuliqujian;
            //获取当前文档及数据库
            //Document acDoc = Application.DocumentManager.MdiActiveDocument;
            //  Database acCurDb = acDoc.Database;

            //启动事务
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                //以读模式打开Block表
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                OpenMode.ForRead) as BlockTable;
                //以写模式打开Block表记录Model空间
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                OpenMode.ForWrite) as BlockTableRecord;
                //创建一个单行文字对象
                using (DBText acText = new DBText())
                {
                    acText.Position = new Point3d(x, y, 0);
                    acText.Height = size;

                    acText.TextString = Text;
                    acText.ColorIndex = colorindex;
                    //文字反向   acText.IsMirroredInX = false;
                    acText.Rotation = hudu;
                    acBlkTblRec.AppendEntity(acText);
                    acTrans.AddNewlyCreatedDBObject(acText, true);


                    //释放DBObject对象
                }
                //保存修改，关闭事务
                acTrans.Commit();
            }
        }

        /// <summary>
        /// 添加矩形面域
        /// </summary>
        /// <param name="color"></param>
        /// <param name="B_x"></param>
        /// <param name="B_y"></param>
        /// <param name="width"></param>
        /// <param name="length"></param>
        /// <param name="acDoc"></param>
        /// <param name="acCurDb"></param>
        public static void AddQuYu(int ColorIndex, double B_x, double B_y, double width, double length,string Filltype, Document acDoc, Database acCurDb)
        {
            //启动事务
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                //以读打开Block表
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                OpenMode.ForRead) as BlockTable;
                //以写方式打开Block表的Model空间记录
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                OpenMode.ForWrite) as BlockTableRecord;
                //在Model空间创建四边形实体（矩形）

                //Create a polyline with two segments (3 points)
                using (Autodesk.AutoCAD.DatabaseServices.Polyline acPoly = new Autodesk.AutoCAD.DatabaseServices.Polyline()) 
                  
                  {
                  
                    acPoly.AddVertexAt(0, new Point2d(B_x, B_y), 0, 0, 0);
                    acPoly.AddVertexAt(1, new Point2d(B_x + length, B_y), 0, 0, 0);
                    acPoly.AddVertexAt(2, new Point2d(B_x + length, B_y - width), 0, 0, 0);
                    acPoly.AddVertexAt(3, new Point2d(B_x, B_y - width), 0, 0, 0);
                    acPoly.AddVertexAt(4, new Point2d(B_x, B_y), 0, 0, 0);
                    //将新对象添加到块表记录并登记事务
                    acBlkTblRec.AppendEntity(acPoly);
                    acPoly.ColorIndex = ColorIndex;

                      acTrans.AddNewlyCreatedDBObject(acPoly, true);
                   

                          // Adds the ac2DSolidBow to an object id array
                          ObjectIdCollection acObjIdColl = new ObjectIdCollection();
                    acObjIdColl.Add(acPoly.ObjectId);
                   AddExtDict( acPoly);
                    
                    //-------------------------------
                    // 创建填充对象
                    //-------------------------------
                    // Create the hatch object and append it to the block table record
                    using (Hatch acHatch = new Hatch())
                    {
                        acBlkTblRec.AppendEntity(acHatch);
                        acTrans.AddNewlyCreatedDBObject(acHatch, true);

                        // Set the properties of the hatch object
                        // Associative must be set after the hatch object is appended to the 
                        // block table record and before AppendLoop
                        acHatch.PatternScale =6;
                        acHatch.SetHatchPattern(HatchPatternType.PreDefined, Filltype);
                      
                        acHatch.Associative = true;
                        acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
                        acHatch.EvaluateHatch(true);
                        acHatch.ColorIndex = ColorIndex;
                        AddExtDict(acHatch);

                    }

                    //释放DBObject对象
                }

                //提交事务，将新对象保存到数据库
                acTrans.Commit();
            }
        }
        /// <summary>
        /// 添加圆面域
        /// </summary>
        public static void AddCircleQuyu(int ColorIndex, double B_x, double B_y,double R, string Filltype, Document acDoc, Database acCurDb)
        {
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                //以读打开Block表
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                OpenMode.ForRead) as BlockTable;
                //以写方式打开Block表的Model空间记录
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                OpenMode.ForWrite) as BlockTableRecord;

                //创建一个圆作为填充的封闭边界
                using (Circle acCirc = new Circle())
                {
                    acCirc.Center = new Point3d(B_x, B_y, 0);
                    acCirc.Radius = R;
                    acCirc.ColorIndex = ColorIndex;
                    //将圆添加到块表记录和事务
                    acBlkTblRec.AppendEntity(acCirc);
                    acTrans.AddNewlyCreatedDBObject(acCirc, true);
                    //将圆的ObjectId添加到对象id数组
                    ObjectIdCollection acObjIdColl = new ObjectIdCollection();
                    acObjIdColl.Add(acCirc.ObjectId);
                    //创建填充对象并添加到块表记录，然后设置其属性
                   
                     using (Hatch acHatch = new Hatch())
                    {
                        acBlkTblRec.AppendEntity(acHatch);
                        acTrans.AddNewlyCreatedDBObject(acHatch, true);

                        // Set the properties of the hatch object
                        // Associative must be set after the hatch object is appended to the 
                        // block table record and before AppendLoop
                        acHatch.PatternScale = 6;
                        acHatch.SetHatchPattern(HatchPatternType.PreDefined, Filltype);

                        acHatch.Associative = true;
                        acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
                        acHatch.EvaluateHatch(true);
                        acHatch.ColorIndex = ColorIndex;
                    }
                }
                //保存新对象到数据库
                acTrans.Commit();
            }

        }

   
        /// <summary>
        /// 设置图例 当是圆时 width=0 length=R
        /// </summary>
        /// <param name="ColorIndex"></param>
        /// <param name="B_x"></param>
        /// <param name="B_y"></param>
        /// <param name="width"></param>
        /// <param name="length"></param>
        /// <param name="Wenzi"></param>
        /// <param name="acDoc"></param>
        /// <param name="acCurDb"></param>
        public static void AddTuli(int ColorIndex, double B_x, double B_y, double width, string Filltype,bool IFCircle, double length,string Wenzi, Document acDoc, Database acCurDb)
        {
            if (IFCircle== true)
            {
                WxxTU.AddCircleQuyu(ColorIndex, B_x, B_y, length, Filltype, acDoc, acCurDb);
            }
            else
            {
                WxxTU.AddQuYu(ColorIndex, B_x, B_y, width, length, Filltype, acDoc, acCurDb);
            }
            WxxTU.ObliqueText(Wenzi, 0, ColorIndex, B_x+ 3, B_y - 3, 2, acDoc, acCurDb);
        }

    
        /// <summary>
        /// 桩号Double转换为string
        /// </summary>
        /// <param name="StationDouble"></param>
        /// <returns></returns>
        public static String Station_Double2Str(double StationDouble)
        {
            double x = Math.Floor(StationDouble / 1000);
            double y = StationDouble % 1000;
            string Y = y.ToString("000.0");
            string Station = "K" + x + "+" + Y;
            return Station;
        }
        /// 命令行中显示字符
        /// </summary>
        /// <param name="message"></param>
        public static void Message(string message,Document doc)
        {
             doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            ed.WriteMessage(message);
        }

        /// <summary>
        /// Reverses the order of the X and Y properties of a Point2d.
        /// </summary>
        /// <param name="flip">Boolean indicating whether to reverse or not.</param>
        /// <returns>The original Point2d or the reversed version.</returns>
        public static Point2d Swap(this Point2d pt, bool flip = true)

        {

            return flip ? new Point2d(pt.Y, pt.X) : pt;

        }

        /// <summary>
        /// Pads a Point2d with a zero Z value, returning a Point3d.
        /// </summary>
        /// <param name="pt">The Point2d to pad.</param>
        /// <returns>The padded Point3d.</returns>
        public static Point3d Pad(this Point2d pt)

        {

            return new Point3d(pt.X, pt.Y, 0);

        }
        /// <summary>
        /// Strips a Point3d down to a Point2d by simply ignoring the Z ordinate.
        /// </summary>
        /// <param name="pt">The Point3d to strip.</param>
        /// <returns>The stripped Point2d.</returns>
        public static Point2d Strip(this Point3d pt)

        {

            return new Point2d(pt.X, pt.Y);

        }
        
        /// <summary>
        /// Creates a layout with the specified name and optionally makes it current.
        /// </summary>
        /// <param name="name">The name of the viewport.</param>
        /// <param name="select">Whether to select it.</param>
        /// <returns>The ObjectId of the newly created viewport.</returns>

        public static ObjectId CreateAndMakeLayoutCurrent(this LayoutManager lm, string name, bool select = true)

        {

            // First try to get the layout



            var id = lm.GetLayoutId(name);



            // If it doesn't exist, we create it



            if (!id.IsValid)

            {

                id = lm.CreateLayout(name);
               

            }



            // And finally we select it



            if (select)

            {

                lm.CurrentLayout = name;

            }



            return id;

        }
        /// <summary>
        /// Applies an action to the specified viewport from this layout.
        /// Creates a new viewport if none is found withthat number.
        /// </summary>
        /// <param name="tr">The transaction to use to open the viewports.</param>
        /// <param name="vpNum">The number of the target viewport.</param>
        /// <param name="f">The action to apply to each of the viewports.</param>

        public static void ApplyToViewport( this Layout lay, Transaction tr, int vpNum, Action<Autodesk.AutoCAD.DatabaseServices.Viewport> f
        )
        {
            var vpIds = lay.GetViewports();
            Autodesk.AutoCAD.DatabaseServices.Viewport vp = null;
         foreach (ObjectId vpId in vpIds)
         {

                var vp2 = tr.GetObject(vpId, OpenMode.ForWrite) as Autodesk.AutoCAD.DatabaseServices.Viewport;
                //&& vp2.Number == vpNum
                if (vp2 != null && vp2.Number == vpNum)

                {

                    // We have found our viewport, so call the action
                    //     vp = vp2;
                    //删除视口
                    // vp2.UpgradeOpen();
                    // vp2.Erase();
                    vp2.GridOn = false;
                    vp2.Visible = false;
                    vp2.Height = 0.001;
                    break;

                }

         }
            


            if (vp == null)

            {

                // We have not found our viewport, so create one



                var btr =

                  (BlockTableRecord)tr.GetObject(

                    lay.BlockTableRecordId, OpenMode.ForWrite

                  );



                vp = new Autodesk.AutoCAD.DatabaseServices.Viewport();



                // Add it to the database



                btr.AppendEntity(vp);

                tr.AddNewlyCreatedDBObject(vp, true);



                // Turn it - and its grid - on


                
                vp.On =true;

                vp.GridOn = false;

                
            }



            // Finally we call our function on it



            f(vp);

        }
        /// <summary>
        /// Apply plot settings to the provided layout.
        /// </summary>
        /// <param name="pageSize">The canonical media name for our page size.</param>
        /// <param name="styleSheet">The pen settings file (ctb or stb).</param>
        /// <param name="devices">The name of the output device.</param>
        public static void SetPlotSettings(this Layout lay, string pageSize, string styleSheet, string device )

        {

            using (var ps = new PlotSettings(lay.ModelType))

            {

                ps.CopyFrom(lay);



                var psv = PlotSettingsValidator.Current;



                // Set the device



                var devs = psv.GetPlotDeviceList();

                if (devs.Contains(device))

                {

                    psv.SetPlotConfigurationName(ps, device, null);

                    psv.RefreshLists(ps);

                }



                // Set the media name/size



                var mns = psv.GetCanonicalMediaNameList(ps);

                if (mns.Contains(pageSize))

                {

                    psv.SetCanonicalMediaName(ps, pageSize);

                }



                // Set the pen settings



                var ssl = psv.GetPlotStyleSheetList();

                if (ssl.Contains(styleSheet))

                {

                    psv.SetCurrentStyleSheet(ps, styleSheet);

                }



                // Copy the PlotSettings data back to the Layout



                var upgraded = false;

                if (!lay.IsWriteEnabled)

                {

                    lay.UpgradeOpen();

                    upgraded = true;

                }



                lay.CopyFrom(ps);



                if (upgraded)

                {

                    lay.DowngradeOpen();

                }

            }

        }
        /// <summary>
        /// Determine the maximum possible size for this layout.
        /// </summary>
        /// <returns>The maximum extents of the viewport on this layout.</returns>
        public static Extents2d GetMaximumExtents(this Layout lay)

        {

            // If the drawing template is imperial, we need to divide by

            // 1" in mm (25.4)



            var div = lay.PlotPaperUnits == PlotPaperUnit.Inches ? 25.4 : 1.0;



            // We need to flip the axes if the plot is rotated by 90 or 270 deg



            var doIt =

              lay.PlotRotation == PlotRotation.Degrees090 ||

              lay.PlotRotation == PlotRotation.Degrees270;



            // Get the extents in the correct units and orientation



            var min = lay.PlotPaperMargins.MinPoint.Swap(doIt) / div;

            var max =

              (lay.PlotPaperSize.Swap(doIt) -

               lay.PlotPaperMargins.MaxPoint.Swap(doIt).GetAsVector()) / div;



            return new Extents2d(min, max);

        }
        /// <summary>
        /// Sets the size of the viewport according to the provided extents.
        /// </summary>
        /// <param name="ext">The extents of the viewport on the page.</param>
        /// <param name="fac">Optional factor to provide padding.</param>
       
        public static void ResizeViewport( this Autodesk.AutoCAD.DatabaseServices.Viewport vp, Extents2d ext, double fac ,
            int num,int i)

        {
            
            vp.CenterPoint = new Point3d(ext.MaxPoint.X * fac / 2+1.2, ext.MaxPoint.Y * fac / 3/2*num+1, 0);
          vp.Width = ext.MaxPoint.X * fac+2;
           
            //- ext.MinPoint.X)- ext.MinPoint.Y)
            vp.Height = ext.MaxPoint.Y  * fac/3;
            double zuobiao =175 * (2*i+1);
            // vp.ViewTarget = new Point3d(zuobiao, 10,0);
            vp.ViewCenter = new Point2d(zuobiao, 5);
           //   vp.StandardScale = StandardScaleType.Scale1To10;
            vp.ViewHeight = 70;
       
            //网格
            vp.GridOn = false;
        

        }
        /// <summary>
        /// Sets the view in a viewport to contain the specified model extents.
        /// </summary>
        /// <param name="ext">The extents of the content to fit the viewport.</param>
        /// <param name="fac">Optional factor to provide padding.</param>

        public static void FitContentToViewport( this Autodesk.AutoCAD.DatabaseServices.Viewport vp, Extents3d ext, double fac = 1.0 )

        {

            // Let's zoom to just larger than the extents



           /* vp.ViewCenter =

              (ext.MinPoint + ((ext.MaxPoint - ext.MinPoint) * 0.5)).Strip();*/



            // Get the dimensions of our view from the database extents



            var hgt = ext.MaxPoint.Y - ext.MinPoint.Y;

            var wid = ext.MaxPoint.X - ext.MinPoint.X;



            // We'll compare with the aspect ratio of the viewport itself

            // (which is derived from the page size)



            var aspect = vp.Width / vp.Height;



            // If our content is wider than the aspect ratio, make sure we

            // set the proposed height to be larger to accommodate the

            // content



            if (wid / hgt > aspect)

            {

                hgt = wid / aspect;

            }



            // Set the height so we're exactly at the extents



            vp.ViewHeight = hgt;



            // Set a custom scale to zoom out slightly (could also

            // vp.ViewHeight *= 1.1, for instance)



            vp.CustomScale *= fac;

        }
        /// <summary>
        /// 删除所有布局
        /// </summary>
        public static void EraseAllLayouts()
        {

            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {

                // ACAD_LAYOUT dictionary.
                DBDictionary layoutDict = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead) as DBDictionary;

                // Iterate dictionary entries.
                foreach (DBDictionaryEntry de in layoutDict)
                {
                    string layoutName = de.Key;
                    Console.WriteLine(layoutName);
                    if (layoutName != "Model")
                    {
                        LayoutManager.Current.DeleteLayout(layoutName); // Delete layout.
                    }
                }
                tr.Commit();
            }

            ed.Regen();   // Updates AutoCAD GUI to relect changes.
        }
        /// <summary>
        /// 创建一条直线
        /// </summary>
        public static void Createline(BlockTableRecord acBlkTblRec, Transaction  tr,double B_x,double E_x,
            double B_Y,double E_Y)
        {
            using (Line acLine = new Line(new Point3d(B_x, B_Y, 0),
                   new Point3d(E_x, E_Y, 0)))
            {
                //将新对象添加到块表记录和事务
                acBlkTblRec.AppendEntity(acLine);
                tr.AddNewlyCreatedDBObject(acLine, true);
                //释放DBObject对象
            }
        }
       /// <summary>
       /// 创建布局和视口
       /// 
       /// </summary>
        public static void CreateLayout(double[] B_Sta, Document doc, Database db)
        {
          //  var doc = Application.DocumentManager.MdiActiveDocument;
           
            if (doc == null)
                return;
           // var db = doc.Database;
            var ed = doc.Editor;
            var ext = new Extents2d();
            int k = 0;// 100000 / 2600 / 3
            int zhushi = 0;
            //求布局数
            double Layoutnum = (B_Sta[B_Sta.Length - 1] - B_Sta[0])/350/3;
            for (int j = 0; j < Layoutnum+1; j++)
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    //以读模式打开块表
                    BlockTable acBlkTbl;
                    acBlkTbl = tr.GetObject(db.BlockTableId,
                    OpenMode.ForRead) as BlockTable;
                    //以写模式打开块表记录Paper空间
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = tr.GetObject(acBlkTbl[BlockTableRecord.PaperSpace],
                    OpenMode.ForWrite) as BlockTableRecord;
                    //切换到Paper空间布局
                    Application.SetSystemVariable("TILEMODE", 0);
                    doc.Editor.SwitchToPaperSpace();
                    // Create and select a new layout tab
                    var id = LayoutManager.Current.CreateAndMakeLayoutCurrent("病害处理" + j);
                    // Open the created layout
                    var lay = (Layout)tr.GetObject(id, OpenMode.ForWrite);
                    // Make some settings on the layout and get its extents
                    /*     lay.SetPlotSettings(
                         //  "ISO_full_bleed_2A0_(1189.00_x_1682.00_MM)", // Try this big boy!
                           "ANSI_B_(11.00_x_17.00_Inches)",
                           "monochrome.ctb",
                           "DWF6 ePlot.pc3");*/
                    //从布局中获取PlotInfo
                    PlotInfo acPlInfo = new PlotInfo();
                    acPlInfo.Layout = lay.ObjectId;
                    //复制布局中的PlotSettings
                    PlotSettings acPlSet = new PlotSettings(lay.ModelType);
                    acPlSet.CopyFrom(lay);
                    //更新PlotSettings对象
                    PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;
                    //设置打印区域
                    acPlSetVdr.SetPlotType(acPlSet,
                    Autodesk.AutoCAD.DatabaseServices.PlotType.Extents);
                    //设置打印比例
                    acPlSetVdr.SetUseStandardScale(acPlSet, true);
                    acPlSetVdr.SetStdScaleType(acPlSet, StdScaleType.ScaleToFit);
                    //居中打印
                    acPlSetVdr.SetPlotCentered(acPlSet, true);
                    ext = lay.GetMaximumExtents();
                    //绘制边框
                    //Create a polyline with two segments (3 points)
                    using (Autodesk.AutoCAD.DatabaseServices.Polyline acPoly = new Autodesk.AutoCAD.DatabaseServices.Polyline())
                        
                    {
                        acPoly.AddVertexAt(0, new Point2d(0.2,0.7), 0, 0, 0);
                        acPoly.AddVertexAt(1, new Point2d(0.2,ext.MaxPoint.Y-0.3), 0, 0, 0);
                        acPoly.AddVertexAt(2, new Point2d(ext.MaxPoint.X-0.13, ext.MaxPoint.Y-0.3), 0, 0, 0);
                        acPoly.AddVertexAt(3, new Point2d(ext.MaxPoint.X-0.13, 0.7), 0, 0, 0);
                        acPoly.AddVertexAt(4, new Point2d(0.2, 0.7), 0, 0, 0);
                        acPoly.LineWeight = LineWeight.LineWeight200;
                        //将新对象添加到块表记录和事务
                        acBlkTblRec.AppendEntity(acPoly);
                        tr.AddNewlyCreatedDBObject(acPoly, true);
                    }
                    //创建下方直线分隔
                    double[] Xiafangzuobiao = new double[] { 1.5,4,7,7.5,8.1,8.6,9.2,9.7,10.3,10.8};
                    for(int i=0;i< Xiafangzuobiao.Length;i++)
                    {
                        WxxTU.Createline(acBlkTblRec, tr, Xiafangzuobiao[i], Xiafangzuobiao[i], 0.7, 1);
                    }
                    //创建下方需要填写的文字
                    string[] Xiafangwenzi = new string[] { "路面病害处治设计图", "设计", "复核", "审核", "图号" };
                    double[] Xiafangwenzizuobiao = new double[] {4.5,7, 8.15,9.2,10.4};
                    for(int i=0;i< Xiafangwenzi.Length;i++)
                    {
                        
                        //创建一个单行文字对象
                        using (DBText acText = new DBText())
                        {
                            acText.Position = new Point3d(Xiafangwenzizuobiao[i], 0.8, 0);
                            acText.Height = 0.15;

                            acText.TextString = Xiafangwenzi[i];
                            acText.ColorIndex = 0;
                            //文字反向   acText.IsMirroredInX = false;
                            acText.Rotation = 0;
                            acBlkTblRec.AppendEntity(acText);
                            tr.AddNewlyCreatedDBObject(acText, true);


                            //释放DBObject对象
                        }
                    }
                    //创建上方直线分割
                    double[] Shangfangzuobiao = new double[] {  10, 10.8 };
                    for (int i = 0; i < Shangfangzuobiao.Length; i++)
                    {
                        WxxTU.Createline(acBlkTblRec, tr, Shangfangzuobiao[i], Shangfangzuobiao[i], ext.MaxPoint.Y - 0.3, ext.MaxPoint.Y - 0.65);
                    }
                    //创建上方需要填写的文字
                    string[] Shangfangwenzi = new string[] { "第" + j  + "页", "共" +  Math.Truncate( Layoutnum) + "页" };
                    for (int i = 0; i < Shangfangwenzi.Length; i++)
                    {

                        //创建一个单行文字对象
                        using (DBText acText = new DBText())
                        {
                            acText.Position = new Point3d(Shangfangzuobiao[i], ext.MaxPoint.Y - 0.5, 0);
                            acText.Height = 0.15;

                            acText.TextString = Shangfangwenzi[i];
                            acText.ColorIndex = 0;
                            //文字反向   acText.IsMirroredInX = false;
                            acText.Rotation = 0;
                            acBlkTblRec.AppendEntity(acText);
                            tr.AddNewlyCreatedDBObject(acText, true);


                            //释放DBObject对象
                        }
                    }
                    //创建桩号起终点注释
                    double[] Sta_Y = new double[] { 2.75, 5, 7.1 };
                    for (int i = 0; i < 3; i++)
                    {
                       
                        using (DBText acText = new DBText())
                        {
                            acText.Position = new Point3d(9, Sta_Y[i], 0);
                            acText.Height = 0.15;

                            acText.TextString = WxxTU.Station_Double2Str(B_Sta[0] + 350 * (zhushi-3)) + "-" + WxxTU.Station_Double2Str(B_Sta[0] + 350 * (zhushi-2));
                            acText.ColorIndex = 0;
                            //文字反向   acText.IsMirroredInX = false;
                            acText.Rotation = 0;
                            acBlkTblRec.AppendEntity(acText);
                            tr.AddNewlyCreatedDBObject(acText, true);
                            zhushi++;

                            //释放DBObject对象
                        }
                    }

                        //创建多个视口
                        for (int i = 0; i < 3; i++)
                    {
                        
                        lay.ApplyToViewport(
                     tr, i,
                     vp =>
                     {

                         // Size the viewport according to the extents calculated when
                         // we set the PlotSettings (device, page size, etc.)
                         // Use the standard 10% margin around the viewport
                         // (found by measuring pixels on screenshots of Layout1, etc.)
                        
                         vp.ResizeViewport(ext, 0.8, 2 * i + 1, k);
                         k++;
                         /* // Adjust the view so that the model contents fit
                           if (ValidDbExtents(db.Extmin, db.Extmax))
                          {
                              vp.FitContentToViewport(new Extents3d(db.Extmin, db.Extmax),0.8);
                          }

                          // Finally we lock the view to prevent meddling*/

                        
                         vp.Locked = true;
                      

                     }
                   );
                    }
                    //  k++;

                    // Commit the transaction
                    tr.Commit();

                }
            }
            // Zoom so that we can see our new layout, again with a little padding
            /*    ed.Command("_.ZOOM", "_E");
                ed.Command("_.ZOOM", ".7X");
                ed.Regen();*/
        }

        /// <summary>
        /// 写扩展词典
        /// </summary
        public static void AddExtDict(DBObject obj)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
           /* PromptEntityOptions peo = new PromptEntityOptions("请选择实体：");
            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status == PromptStatus.OK)
            {*/
                Transaction trans = doc.TransactionManager.StartTransaction();
             //   DBObject obj = trans.GetObject(per.ObjectId, OpenMode.ForWrite);
                AddRegAppTableRecord("扩展数据测试");
                ResultBuffer rb = new ResultBuffer();
                rb.Add(new TypedValue(1001, "扩展数据测试"));
                rb.Add(new TypedValue(1000, "我是扩展数据，大家好！"));
                rb.Add(new TypedValue(1000, "看到我你就测试成功了！"));
                obj.XData = rb;
                rb.Dispose();
                trans.Commit();
                trans.Dispose();
           // }
        }
        private static void AddRegAppTableRecord(string regAppName)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            Transaction trans = doc.TransactionManager.StartTransaction();
            RegAppTable rat = (RegAppTable)trans.GetObject(db.RegAppTableId, OpenMode.ForRead, false);
            if (!rat.Has(regAppName))
            {
                rat.UpgradeOpen();
                RegAppTableRecord ratr = new RegAppTableRecord();
                ratr.Name = regAppName;
                rat.Add(ratr);
                trans.AddNewlyCreatedDBObject(ratr, true);
            }
            trans.Commit();
            trans.Dispose();
        }
    }
}


