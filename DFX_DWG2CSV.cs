 public void extractSettingsFromDXE()

        {

        // Add the AcDx.dll reference from the inc folder

             Document doc = Application.DocumentManager.MdiActiveDocument;

            Editor ed = doc.Editor;

            Database db = doc.Database;

            StringBuilder fileContent = new StringBuilder();

            const string dxePath = @"C:\Mechanical_Multileaders.dxe";

            if (System.IO.File.Exists(dxePath))

            {

 

            /*Load DxE from disk*/

           IDxExtractionSettings extractionSettings = DxExtractionSettings.FromFile(dxePath);

 

 

            /*Retrieve Information about File Structure*/

           DxFileList files = extractionSettings.DrawingDataExtractor.Settings.DrawingList as DxFileList;

           IDxFileReference[] fileRefereces = files.Files;

           foreach (DxFileReference dwgFile in fileRefereces)

               {

               ed.WriteMessage("\nDrawingFile :{0}", dwgFile);

               }

 

           IDxFileReference[] xrefFiles = files.XrefFiles;

           foreach (DxFileReference xref in xrefFiles)

               {

               ed.WriteMessage("\nXref DrawingFile :{0}", xref);

              }

            /*Write data to CSV*/

           if (extractionSettings.DrawingDataExtractor.ExtractData(dxePath))

               {

               System.Data.DataTable dt = extractionSettings.DrawingDataExtractor.ExtractedData;

 

               foreach (var col in dt.Columns)

                   {

                   fileContent.Append(col.ToString() + ",");

                   }

 

               fileContent.Replace(",", System.Environment.NewLine, fileContent.Length - 1, 1);

                foreach (DataRow dr in dt.Rows)

                   {

                    foreach (var column in dr.ItemArray)

                       {

                       fileContent.Append("\"" + column.ToString() + "\",");

                       }

 

                   fileContent.Replace(",", System.Environment.NewLine, fileContent.Length - 1, 1);

                   }

               if(File.Exists(@"C:\MLeaders.csv"))

                   File.Delete(@"C:\MLeaders.csv");

               System.IO.File.WriteAllText(@"C:\MLeaders.csv", fileContent.ToString());

               }

 

            /*Some Other Information*/

           //IDxOutputSettings outPutSettings = extractionSettings.OutputSettings;

           //AdoOutput.OutputType outPutType =  outPutSettings.FileOutputType;

 

           //DxOuputFlags oFlags = outPutSettings.OuputFlags;

           //IDxReport report  =  extractionSettings.Report;

 

            }

        }