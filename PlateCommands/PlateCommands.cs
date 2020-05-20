using Autodesk.AutoCAD.Runtime;
using System;

using PlateData;
using Tools;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.GraphicsInterface;

namespace PlateCommands
{
    public class PlateCommands
    {
        // properties
        // - colors
        static int CONTOUR = 0;                 // contour color
        static int DIMENSION = 1;               // dimension color
        static int HIDDEN = 2;                  // hidden line color
        static int CENTER = 3;                  // center line color

        static int DIMNO = 0;                   // no dimension 
        static int DIMHOR = 0x01;               // add horizontal dimension
        static int DIMVER = 0x02;               // add vertical dimension
        
        // create 1rst command - to draw 2d drawing
        [CommandMethod("CreatePlate2d")]
        public void CreatePlate2d()
        {
            // create Plate object
            Log.Append("> create plate...");
            Plate plate = new Plate();

            // read Plate data from xls sheet
            Log.Append("> load plate data...");
            if (!plate.Load()) return;

            // check input data
            int nErr = plate.Check();
            if (nErr > 0)
            {
                Log.Append(string.Format("*** error: invalid plate data! {0} errors found.", nErr));
                return;
            }

            // Load linetypes from the linetype file
            //...\AutoCad\AutoCAD 2016\UserDataCache\en-us\Support\acad.lin
            AcTrans.LoadlineType("HIDDEN");       // VERDECKT in German 
            AcTrans.LoadlineType("DASHDOT");
            AcTrans.LoadlineType("CENTER");

            // draw the top view
            Point3d p0 = new Point3d(0.0, 0.0, 0.0);
            DrawProfile(plate, p0, DIMNO);

            // draw the side view
            Point3d p1 = new Point3d(plate.dW + plate.dViewSp, 0.0, 0.0);
            DrawSideRectangle(plate, p1, plate.dT, plate.dL, DIMVER);

            // draw the front view
            Point3d p2 = new Point3d(plate.dW, -(plate.dT + plate.dViewSp), 0.0);
            DrawFrontRectangle(plate, p2, -2*plate.dW, plate.dT, DIMHOR);

        }

        // create 2nd command - to draw 3d drawing
        [CommandMethod("CreateSolidPlate")]
        public void CreateSolidPlate()
        {
            // create Plate object
            Log.Append("> create plate...");
            Plate plate = new Plate();

            // read Plate data from xls sheet
            Log.Append("> load plate data...");
            if (!plate.Load()) return;

            // check input data
            int nErr = plate.Check();
            if (nErr > 0)
            {
                Log.Append(string.Format("*** error: invalid plate data! {0} errors found.", nErr));
                return;
            }

            // join lines or create region
            // draw the profile without hole
            Point3d p0 = new Point3d(0.0, 0.0, 0.0);
            DrawSolidProfile(plate, p0, DIMNO);

        }

        // draw the lower three lines
        // nFlag: dimension control
        private void DrawProfile(Plate pl, Point3d p0, int nFlag)
        {
            
            Point3d[] pt = new Point3d[26];  // geometry helpers
            Line[] ln = new Line[26];        // database objects

            // create points
            pt[0] = p0;
            pt[1] = new Point3d(p0.X + (pl.dW - pl.R1), p0.Y, p0.Z);
            pt[2] = new Point3d(p0.X + pl.dW, pl.R1, p0.Z);
            pt[3] = new Point3d(pt[1].X, pt[2].Y, p0.Z);
            pt[4] = new Point3d(pt[3].X, pt[3].Y+pl.L9, p0.Z);
            pt[5] = new Point3d(pt[4].X + pl.L2, pt[4].Y + pl.L8, p0.Z);
            pt[6] = new Point3d(pt[5].X, pt[5].Y + pl.L7, p0.Z);
            pt[7] = new Point3d(pt[4].X, pt[6].Y + pl.L6, p0.Z);
            pt[8] = new Point3d(pt[7].X, pt[7].Y + pl.L5, p0.Z);
            pt[9] = new Point3d(p0.X + pl.L3, pt[8].Y + pl.L4, p0.Z);
            pt[10] = new Point3d(p0.X, pt[9].Y, p0.Z);

            // symmetry - create points on negative x-axis
            int k = 9;
            for (int i = 11; i < 20; i++)
            {    
                pt[i] = new Point3d(-pt[k].X, pt[k].Y, pt[k].Z);
                k--;
            }

            // Draw Center symmetry symbol
            Line lnc = new Line(new Point3d(p0.X, p0.Y - 2*pl.dT, p0.Z), new Point3d(pt[10].X, pt[10].Y + 2*pl.dT, p0.Z));
            lnc.ColorIndex = pl.nColor[CENTER];
            lnc.Linetype = "DASHDOT";
            lnc.LinetypeScale = 50.0;
            AcTrans.Add(lnc);

            // create lines and store them into the database
            for (int i = 0; i < 20; i++)
            {
                if (i != 1  & i != 18)
                {
                    ln[i] = new Line(pt[i], pt[(i + 1)]);
                    ln[i].ColorIndex = pl.nColor[CONTOUR];      // 0:index
                    AcTrans.Add(ln[i]);
                }
                else if (i == 1)
                {
                    double st_angle0 = 270 * (Math.PI / 180);
                    double end_angle0 = 0;

                    Arc a0 = new Arc(pt[3], pl.R1, st_angle0, end_angle0);
                    a0.ColorIndex = pl.nColor[CONTOUR];      // 0:index
                    AcTrans.Add(a0);

                    // create the holes dimension instance
                    RadialDimension radDim = new RadialDimension();
                    radDim.Center = a0.Center;
                    radDim.ChordPoint = new Point3d(a0.Center.X + pl.R1,
                                                    a0.Center.Y, 0.0);
                    radDim.LeaderLength = 10.0;  // !!! not optimal
                    radDim.Dimscale = pl.dDimScale;
                    // set current dimensionstyle
                    radDim.ColorIndex = pl.nColor[DIMENSION];
                    AcTrans.Add(radDim);
                }
                else if (i == 18)
                {
                    double st_angle1  = 180 * (Math.PI / 180);
                    double end_angle1 = 270 * (Math.PI / 180);

                    Arc a1 = new Arc(pt[17], pl.R1, st_angle1, end_angle1);
                    a1.ColorIndex = pl.nColor[CONTOUR];      // 0:index
                    AcTrans.Add(a1);
                }
                else
                {
                    continue;
                }
            }

            // draw points for centre hole
            pt[20] = new Point3d(p0.X + pl.R2, pt[3].Y + pl.R2, 0.0);
            pt[21] = new Point3d(pt[20].X, pt[7].Y - pl.R2, 0.0);
            pt[22] = new Point3d(-pt[20].X, pt[7].Y - pl.R2, 0.0);
            pt[23] = new Point3d(-(p0.X + pl.R2), pt[3].Y + pl.R2, 0.0);
            pt[24] = new Point3d(p0.X, pt[20].Y, 0.0);
            pt[25] = new Point3d(p0.X, pt[21].Y, 0.0);

            // draw lines for centre hole
            for (int i = 20; i < 24; i++)
            {
                if (i != 21 & i != 23)
                {
                    ln[i] = new Line(pt[i], pt[(i + 1)]);
                    ln[i].ColorIndex = pl.nColor[CONTOUR];      // 0:index
                    AcTrans.Add(ln[i]);
                }
                else if (i == 21)
                {
                    double st_angle2 = 0;
                    double end_angle2 = 180 * (Math.PI / 180);

                    Arc a2 = new Arc(pt[25], pl.R2, st_angle2, end_angle2);
                    a2.ColorIndex = pl.nColor[CONTOUR];      // 0:index
                    AcTrans.Add(a2);

                    // create the holes dimension instance
                    RadialDimension radDim = new RadialDimension();
                    radDim.Center = a2.Center;
                    radDim.ChordPoint = new Point3d(a2.Center.X + pl.R2,
                                                    a2.Center.Y, 0.0);
                    radDim.LeaderLength = 10.0;  // !!! not optimal
                    radDim.Dimscale = pl.dDimScale;
                    // set current dimensionstyle
                    radDim.ColorIndex = pl.nColor[DIMENSION];
                    AcTrans.Add(radDim);

                }
                else if (i == 23 )
                {
                    double st_angle3 = 180 * (Math.PI / 180);
                    double end_angle3 = 0;

                    Arc a3 = new Arc(pt[24], pl.R2, st_angle3, end_angle3);
                    a3.ColorIndex = pl.nColor[CONTOUR];      // 0:index
                    AcTrans.Add(a3);
                }
                else
                {
                    continue;
                }
                    
            }

        }

        // draw the 3-D drawing
        // nFlag: dimension control
        private void DrawSolidProfile(Plate pl, Point3d p0, int nFlag)
        {

            // Create Container for the curve, the region (face)
            DBObjectCollection curves = new DBObjectCollection();
            // - Collection for the regions
            DBObjectCollection regions = new DBObjectCollection();

            // Solid3d objects
            Solid3d plateSolid = new Solid3d();
            Solid3d rectHole = new Solid3d();              // Rectangular Hole in the Center

            Point3d[] pt = new Point3d[26];                // geometry helpers
            Line[] ln = new Line[26];                      // database objects

            // create points
            pt[0] = p0;
            pt[1] = new Point3d(p0.X + (pl.dW - pl.R1), p0.Y, p0.Z);
            pt[2] = new Point3d(p0.X + pl.dW, pl.R1, p0.Z);
            pt[3] = new Point3d(pt[1].X, pt[2].Y, p0.Z);
            pt[4] = new Point3d(pt[3].X, pt[3].Y + pl.L9, p0.Z);
            pt[5] = new Point3d(pt[4].X + pl.L2, pt[4].Y + pl.L8, p0.Z);
            pt[6] = new Point3d(pt[5].X, pt[5].Y + pl.L7, p0.Z);
            pt[7] = new Point3d(pt[4].X, pt[6].Y + pl.L6, p0.Z);
            pt[8] = new Point3d(pt[7].X, pt[7].Y + pl.L5, p0.Z);
            pt[9] = new Point3d(p0.X + pl.L3, pt[8].Y + pl.L4, p0.Z);
            pt[10] = new Point3d(p0.X, pt[9].Y, p0.Z);

            // symmetry - create points on negative x-axis
            int k = 9;
            for (int i = 11; i < 20; i++)
            {
                pt[i] = new Point3d(-pt[k].X, pt[k].Y, pt[k].Z);
                k--;
            }

            // create lines and store them into the database
            for (int i = 0; i < 20; i++)
            {
                if (i != 1 & i != 18)
                {
                    ln[i] = new Line(pt[i], pt[(i + 1)]);
                    ln[i].ColorIndex = pl.nColor[CONTOUR];      // 0:index
                    curves.Add(ln[i]);
                }
                else if (i == 1)
                {
                    double st_angle0 = 270 * (Math.PI / 180);
                    double end_angle0 = 0;

                    Arc a0 = new Arc(pt[3], pl.R1, st_angle0, end_angle0);
                    a0.ColorIndex = pl.nColor[CONTOUR];      // 0:index
                    curves.Add(a0);
                }
                else if (i == 18)
                {
                    double st_angle1 = 180 * (Math.PI / 180);
                    double end_angle1 = 270 * (Math.PI / 180);

                    Arc a1 = new Arc(pt[17], pl.R1, st_angle1, end_angle1);
                    a1.ColorIndex = pl.nColor[CONTOUR];      // 0:index
                    curves.Add(a1);
                }
                else
                {
                    continue;
                }
            }

            // Create the Region from the curves
            regions = Region.CreateFromCurves(curves);
            plateSolid.ColorIndex = pl.nColor[CONTOUR];

            // Create the plate without the hole
            plateSolid.Extrude((Region)regions[0], pl.dT, 0.0);

            // draw points for centre hole
            curves.Clear();

            pt[20] = new Point3d(p0.X + pl.R2, pt[3].Y + pl.R2, 0.0);
            pt[21] = new Point3d(pt[20].X, pt[7].Y - pl.R2, 0.0);
            pt[22] = new Point3d(-pt[20].X, pt[7].Y - pl.R2, 0.0);
            pt[23] = new Point3d(-(p0.X + pl.R2), pt[3].Y + pl.R2, 0.0);
            pt[24] = new Point3d(p0.X, pt[20].Y, 0.0);
            pt[25] = new Point3d(p0.X, pt[21].Y, 0.0);

            // draw lines for centre hole
            for (int i = 20; i < 24; i++)
            {
                if (i != 21 & i != 23)
                {
                    ln[i] = new Line(pt[i], pt[(i + 1)]);
                    ln[i].ColorIndex = pl.nColor[CONTOUR];      // 0:index
                    curves.Add(ln[i]);
                }
                else if (i == 21)
                {
                    double st_angle2 = 0;
                    double end_angle2 = 180 * (Math.PI / 180);

                    Arc a2 = new Arc(pt[25], pl.R2, st_angle2, end_angle2);
                    a2.ColorIndex = pl.nColor[CONTOUR];      // 0:index
                    curves.Add(a2);
                }
                else if (i == 23)
                {
                    double st_angle3 = 180 * (Math.PI / 180);
                    double end_angle3 = 0;

                    Arc a3 = new Arc(pt[24], pl.R2, st_angle3, end_angle3);
                    a3.ColorIndex = pl.nColor[CONTOUR];      // 0:index
                    curves.Add(a3);
                }
                else
                {
                    continue;
                }

            }

            // Create the Region from the curves (Rectangular Hole)
            regions = Region.CreateFromCurves(curves);
            rectHole.Extrude((Region)regions[0], pl.dT, 0.0);
            rectHole.ColorIndex = pl.nColor[HIDDEN];

            // Substract the Hole Solid from the Rectangle
            plateSolid.BooleanOperation(BooleanOperationType.BoolSubtract, rectHole);

            // Add Solids into the Database
            AcTrans.Add(plateSolid);

        }

        // draw a rectrangle with dimensions
        // nFlag: dimension control
        private void DrawFrontRectangle(Plate pl, Point3d p0,
                                   double dL, double dW, int nFlag)
        {
            Point3d[] pt = new Point3d[16];  // geometry helpers
            Line[] ln = new Line[16];        // database objects

            // create points
            pt[0] = p0;
            pt[1] = new Point3d(p0.X + dL, p0.Y, p0.Z);
            pt[2] = new Point3d(p0.X + dL, p0.Y + dW, p0.Z);
            pt[3] = new Point3d(p0.X, p0.Y + dW, p0.Z);
            //more points
            pt[4] = new Point3d(p0.X - pl.L1, p0.Y, p0.Z);
            pt[5] = new Point3d(p0.X - pl.L1, p0.Y + dW, p0.Z);
            pt[6] = new Point3d(pt[4].X - pl.L2, p0.Y, p0.Z);
            pt[7] = new Point3d(pt[4].X - pl.L2, p0.Y + dW, p0.Z);
            pt[8] = new Point3d(pl.L3, p0.Y, p0.Z);
            pt[9] = new Point3d(pl.L3, p0.Y + dW, p0.Z);
            pt[10] = new Point3d(-pl.L3, p0.Y, p0.Z);
            pt[11] = new Point3d(-pl.L3, p0.Y + dW, p0.Z);
            pt[12] = new Point3d(-(pt[4].X - pl.L2), p0.Y, p0.Z);
            pt[13] = new Point3d(-(pt[4].X - pl.L2), p0.Y + dW, p0.Z);
            pt[14] = new Point3d(-(p0.X - pl.L1), p0.Y, p0.Z);
            pt[15] = new Point3d(-(p0.X - pl.L1), p0.Y + dW, p0.Z);

            // create lines of the main rectangle and store them into the database
            for (int i = 0; i < 4; i++)
            {
                ln[i] = new Line(pt[i], pt[(i + 1) % 4]);
                ln[i].ColorIndex = pl.nColor[CONTOUR];      // 0:index
                AcTrans.Add(ln[i]);
            }

            // create vertical lines
            for (int i = 4; i < 15; i += 2)
            {
                ln[i] = new Line(pt[i], pt[(i + 1)]);
                ln[i].ColorIndex = pl.nColor[HIDDEN];      // 0:index
                AcTrans.Add(ln[i]);
            }

            // Draw the horizontal dimension
            if ((nFlag & DIMHOR) != 0)
            {
                RotatedDimension dim = new RotatedDimension();
                dim.XLine1Point = pt[0];
                dim.XLine2Point = pt[4];
                dim.DimLinePoint = new Point3d(pt[0].X, pt[0].Y - pl.dDimLineSp, pt[0].Z);
                dim.Dimscale = pl.dDimScale;
                dim.ColorIndex = pl.nColor[DIMENSION];
                dim.Dimzin = 12;                                // Suppresses the Decimal Point   
                AcTrans.Add(dim);

                for (int i = 4; i < 14; i += 2)
                {
                    RotatedDimension dim1 = new RotatedDimension();
                    dim1.XLine1Point = pt[i];
                    dim1.XLine2Point = pt[i+2];
                    dim1.DimLinePoint = new Point3d(pt[0].X, pt[0].Y - pl.dDimLineSp, pt[0].Z);
                    dim1.Dimscale = pl.dDimScale;
                    dim1.ColorIndex = pl.nColor[DIMENSION];
                    dim1.Dimzin = 12;                                // Suppresses the Decimal Point   
                    AcTrans.Add(dim1);
                }

                RotatedDimension dim2 = new RotatedDimension();
                dim2.XLine1Point = pt[14];
                dim2.XLine2Point = pt[1];
                dim2.DimLinePoint = new Point3d(pt[0].X, pt[0].Y - pl.dDimLineSp, pt[0].Z);
                dim2.Dimscale = pl.dDimScale;
                dim2.ColorIndex = pl.nColor[DIMENSION];
                dim2.Dimzin = 12;                                // Suppresses the Decimal Point   
                AcTrans.Add(dim2);
            }
            // Draw the vertical dimension
            if ((nFlag & DIMVER) != 0)
            {
                RotatedDimension dim = new RotatedDimension();
                dim.XLine1Point = pt[1];
                dim.XLine2Point = pt[2];
                dim.DimLinePoint = new Point3d(pt[1].X + pl.dDimLineSp, pt[1].Y, pt[0].Z);
                dim.Rotation = Math.PI / 2.0;
                dim.Dimscale = pl.dDimScale;
                dim.ColorIndex = pl.nColor[DIMENSION];
                dim.Dimzin = 12;                                 // Suppresses the Decimal Point   
                AcTrans.Add(dim);
            }
        }

        // draw a rectrangle with dimensions
        // nFlag: dimension control
        private void DrawSideRectangle(Plate pl, Point3d p0,
                                   double dT, double dL, int nFlag)
        {
            Point3d[] pt = new Point3d[16];  // geometry helpers
            Line[] ln = new Line[16];        // database objects

            // create points
            pt[0] = p0;
            pt[1] = new Point3d(p0.X + dT, p0.Y, p0.Z);
            pt[2] = new Point3d(p0.X + dT, p0.Y + dL, p0.Z);
            pt[3] = new Point3d(p0.X, p0.Y + dL, p0.Z);
            
            //more points
            pt[4] = new Point3d(p0.X, p0.Y + pl.R1, p0.Z);
            pt[5] = new Point3d(p0.X + dT, p0.Y + pl.R1, p0.Z);
            pt[6] = new Point3d(p0.X, pt[4].Y + pl.L9, p0.Z);
            pt[7] = new Point3d(p0.X + dT, pt[4].Y + pl.L9, p0.Z);
            pt[8] = new Point3d(p0.X, pt[6].Y + pl.L8, p0.Z);
            pt[9] = new Point3d(p0.X + dT, pt[6].Y + pl.L8, p0.Z);
            pt[10] = new Point3d(p0.X, pt[8].Y + pl.L7, p0.Z);
            pt[11] = new Point3d(p0.X + dT, pt[8].Y + pl.L7, p0.Z);
            pt[12] = new Point3d(p0.X, pt[10].Y + pl.L6, p0.Z);
            pt[13] = new Point3d(p0.X + dT, pt[10].Y + pl.L6, p0.Z);
            pt[14] = new Point3d(p0.X, pt[12].Y + pl.L5, p0.Z);
            pt[15] = new Point3d(p0.X + dT, pt[12].Y + pl.L5, p0.Z);

            // create lines of the main rectangle and store them into the database
            for (int i = 0; i < 4; i++)
            {
                ln[i] = new Line(pt[i], pt[(i + 1) % 4]);
                ln[i].ColorIndex = pl.nColor[CONTOUR];      // 0:index
                AcTrans.Add(ln[i]);
            }

            // create vertical lines
            for (int i = 4; i < 15; i += 2)
            {
                ln[i] = new Line(pt[i], pt[(i + 1)]);
                ln[i].ColorIndex = pl.nColor[HIDDEN];      // 0:index
                AcTrans.Add(ln[i]);
            }

            // Draw the horizontal dimension
            if ((nFlag & DIMHOR) != 0)
            {
                RotatedDimension dim = new RotatedDimension();
                dim.XLine1Point = pt[0];
                dim.XLine2Point = pt[1];
                dim.DimLinePoint = new Point3d(pt[0].X, pt[0].Y - pl.dDimLineSp, pt[0].Z);
                dim.Dimscale = pl.dDimScale;
                dim.ColorIndex = pl.nColor[DIMENSION];
                dim.Dimzin = 12;                                // Suppresses the Decimal Point   
                AcTrans.Add(dim);
            }

            // Draw the vertical dimension
            if ((nFlag & DIMVER) != 0)
            {
                RotatedDimension dim = new RotatedDimension();
                dim.XLine1Point = pt[1];
                dim.XLine2Point = pt[5];
                dim.DimLinePoint = new Point3d(pt[1].X + pl.dDimLineSp, pt[1].Y, pt[0].Z);
                dim.Rotation = Math.PI / 2.0;
                dim.Dimscale = pl.dDimScale;
                dim.ColorIndex = pl.nColor[DIMENSION];
                dim.Dimzin = 12;                                // Suppresses the Decimal Point   
                AcTrans.Add(dim);

                for (int i = 5; i < 15; i += 2)
                {
                    RotatedDimension dim1 = new RotatedDimension();
                    dim1.XLine1Point = pt[i];
                    dim1.XLine2Point = pt[i + 2];
                    dim1.DimLinePoint = new Point3d(pt[1].X + pl.dDimLineSp, pt[1].Y, pt[0].Z);
                    dim1.Rotation = Math.PI / 2.0;
                    dim1.Dimscale = pl.dDimScale;
                    dim1.ColorIndex = pl.nColor[DIMENSION];
                    dim1.Dimzin = 12;                                // Suppresses the Decimal Point   
                    AcTrans.Add(dim1);
                }

                RotatedDimension dim2 = new RotatedDimension();
                dim2.XLine1Point = pt[15];
                dim2.XLine2Point = pt[2];
                dim2.DimLinePoint = new Point3d(pt[1].X + pl.dDimLineSp, pt[1].Y, pt[0].Z);
                dim2.Rotation = Math.PI / 2.0;
                dim2.Dimscale = pl.dDimScale;
                dim2.ColorIndex = pl.nColor[DIMENSION];
                dim2.Dimzin = 12;                                // Suppresses the Decimal Point   
                AcTrans.Add(dim2);

            }

        }

    }

}
