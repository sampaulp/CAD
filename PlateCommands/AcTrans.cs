using System;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;

namespace Tools
{
    // Database interface
    class AcTrans
    {
        // write the AC entity into the AC database
        public static void Add(Entity entity)
        {
            // get database reference
            Database db = Application.DocumentManager.MdiActiveDocument.Database;

            using (Transaction acTrans = db.TransactionManager.StartTransaction())
            {
                Log.Append("> Start Transaction...");

                try
                {
                    // create a block table record
                    BlockTableRecord acBTR = (BlockTableRecord)acTrans.GetObject(
                                db.CurrentSpaceId,      // space id (f.e. model space)
                                OpenMode.ForWrite);

                    // put entity into block table record
                    acBTR.AppendEntity(entity);

                    // put it into the transaction
                    //                                      |true add, false delete
                    acTrans.AddNewlyCreatedDBObject(entity, true);

                    // write data into the database completing the transaction
                    acTrans.Commit();

                    Log.Append("> Transaction successfull!");
                }

                // handle autocad exception
                catch(Autodesk.AutoCAD.Runtime.Exception ex)
                {
                    acTrans.Abort();
                    Log.Append("*** error in transaction!");
                    Log.Append(ex.ToString());
                }

                // clearing
                finally
                {
                    GC.Collect();   // clear the memory
                }
                Log.Append("> completed!");
            }

        }

    }
}
