# Repo-KCIN3D-reconstruction

## For running the prototype, run only  
    in Branch-4
    Repo-KCIN3D-reconstruction/Initial_Prototype_UI/Komatsu_3D.bat on Windows-terminal 
    
    Download/save the folders in Branch-4 and keep the folder structure as it is.
    -- Komatsu_3D.bat - THE ONLY FILE to run on Windows-Terminal 
    -- consol_interactive_ALPHA_3.3.py - the script for DX-1
    -- 3DFObjects.vbs - .NET to communicate extracted datapoints with CAD for visualization 
    ----------- /Inputs/ includes projects' datasets
    ----------- /Outputs
    -----------/Temp
    -----------/src / includes and will be added more executive scripts
    
    Notes: 
    In ./Inputs/ 183D1PLZ_平面図04.csv (Includes the datapoints of Plan-View. Enter the file name without extension when you initiate the python script) and 
    187D1PFZ_縦断図03.csv  (Includes the datapoints of Longitudinal-Cross-Section-View. Enter the file name as the third input when you initiate the python script)
    
    1- If you wish to visualize the results, You need to have AutoCAD or C3D been running in background 
    2- The script automatically draws the road structure in 3D feature-objects 
    3- You need to install these dependencies / requirements :   
            pandas 
            numpy 
            math 
            openpyxl 
            pathlib 
            win32com 
            xlsxwriter 
            
     +++ Tested in Visual Studio with Python 3.6, 3.8, 3.9 and datasets from 3 different road-projects 

## For Demo, use the files in Branch-3
        Repo-KCIN3D-reconstruction/First_Demo_Windows_Based/
        +++ Follow ReadMe there
        +++ Tested in Visual Studio with Python 3.6, 3.8, 3.9


## For replication, use the files in 
        Repo-KCIN3D-reconstruction/4_UI_BEFORE_Shell/

### For the prototype KMCIN1 and feasibility study Feb-March 2021

Move to KMCIN2 Proof of Concept March-May 2021
        On April 24 2021
         New revision 
         this file Repo-KCIN3D-reconstruction/4_UI_BEFORE_Shell/consol_interactive_ALPHA_2.py was added 
         # This python code automatically run VBA in excel ... It should be run in Windows Environment. 
         
         this file Repo-KCIN3D-reconstruction/4_UI_BEFORE_Shell/Demo1_183D1PLZ_平面図04_OUT0.xlsm was added
         # This file contains the VBA for communicating between output-datasets and AutoCAD 
        
        On April 23 2021
         this file Repo-KCIN3D-reconstruction/4_UI_BEFORE_Shell/3D_Objects_CR_Adjusted_for_Komatsu_2.bas was revised 
         this file Repo-KCIN3D-reconstruction/4_UI_BEFORE_Shell/consol_interactiveALPHA_1.1.py was revised 
        
        On Apr 16 2021
         One file is added XLMacro_Python_Rub.py (run VBA from Python)
         One file is added consol_interactiveALPHA_0.py - early stage of development - version ALPHA
         One file is revised DX_Komatsu_Prototype_0-183D1PLZ_平面図04_REV3.pynb
         One file is revised DX_Komatsu_Prototype_0-183D1PLZ_平面図04_REV2.pynb
         One file is added consol_Interactive.ipynb - far early stage of development - version ALPHA
        
        
        On Apr 14 2021
         one File is revised DX_Komatsu_Prototype_0-233D1CSZ_横断図25.ipynb
        
        
        On Apr 7 2021
         One file is revised DX_Komatsu_Prototype_0-183D1PLZ_平面図04_REV1.ipynb
         
         Added 
         
         >>> DX_Komatsu_Prototype_0-233D1CSZ_横断図25.ipynb
         >>> DFX_DWG2CSV.cs
         >>> vs_community__2111406413.1569763062.exe


## Contexts and background 

Our clients usually receive road-construction drawings from some engineering firms (third-parties). These drawings are detailed the construction procedures and the roads, designs and often come in interconnected 2D drawings. The drawings are complex as they contain huge volume of unstructured data either in symbols, or natural languages, or geometrical objects. 

        Challenge-1- The clients need to extract specific construction and design information about the road objects in a 3D coordinates system,

        Challenge-2 - The extracted information should be extremely accurate and be consolidated in a fashion to create the 3D surfaces which represent the roads in the Earth                          Ground (EG) or the natural ground.

For each road-project, the clients usually have four types of drawings / Datasets:

1. Plot-view drawing(s) which show an overview (top) of the roads with all details associated with the roads' surroundings (PVDwg)
2. Profile-view drawing(s) which show the longitudinal cross-sections of the roads. These drawings contain all information about elevations of the roads (LVDwg)
3. Road-Assembly drawings which are detailed the roads and their shoulders and sub-assemblies and in effect they show the transversal cross-sections of the roads in different parts along the roads (RTDwg)
4. Ground Surface elevation datasets which contain the topographic information of the Earth Ground (EG).

 To address the challenges, we shall introduce these two approaches:

## Approach-1 

includes eight steps of which the first four steps are identical to the steps of Approach-2. In this approach, we will try to minimize dependency to C3D functionalities. The outputs of this approach will be the 3D view of the roads-surface (TIN), the 3D roads information in .XML files. The metric that shows the performance of our model/approach will be the distance between the already-available XML data of the roads (as the ground truth) and our model’s generated data outputs. 

1. Annotation and Vectorization: At this step, we will convert all drawings to datapoints with their relevant features, save in csv files. Each project will have its own Master-Folder.
2. Data pre-processing and management: we will use Python to clean, prepare, and manage all the extracted datapoints from the drawings
3. Information Extraction-1: We will use Python to extract desired information (queries) about the roads centerlines and the stations from the PVDwg(s). We will continue to extract complimentary information (elevations) from LVDwg(s) and build a “Data-Frame” that contains 3D information about the roads centerlines. We will save the Data-Frame in a csv file
4. Information Extraction-2: We will use python to extract all information about the road-finishing surfaces (RTDwg) along with their shoulders, subassemblies, and ditches. If the road-finishing surfaces are not available, we will apply the rule (created based on engineering knowledge of Komatsu) and approximate the finishing surface. The information and datapoints are save in different csv files; each tagged with a station number
5. Re-creation of the roads centerlines 3D: we will use VB or C#
6. Translation of the road-transversal-cross-sections 2D: we will use VB or C#
7. Importing EG data: we will use VB.NET or C#.NET
8. Creating 3D surface we will use VB.NET or C#.NET


## Approach-2 

includes ten steps of which the first four steps are identical to the steps of Approach-1. In this approach, we will use C3D functionalities to create engineering entities which may not be useful for Komatsu or the project objectives, though they will make our approach more smoother and perhaps efficient. The outputs of this approach will be the 3D view of the roads-surface (TIN), a 3D corridor, a set of Komatsu’s road-cross-section libraries (which can be reused for other projects), the 3D roads information in .XML files. The metric that shows the performance of our model/approach will be the distance between the already-available XML data of the roads (as the ground truth) and our model’s generated data outputs.   

1. Annotation and Vectorization: At this step, we will convert all drawings to datapoints with their relevant features, save in csv files. Each project will have its own Master-Folder
2. Data pre-processing and management: we will use Python to clean, prepare, and manage all the extracted datapoints from the drawings
3. Information Extraction-1: We will use Python to extract desired information (queries) about the roads centerlines and the stations from the PVDwg(s). We will continue to extract complimentary information (elevations) from LVDwg(s) and build a “Data-Frame” that contains 3D information about the roads centerlines. We will save the Data-Frame in a csv file
4. Information Extraction-2: We will use python to extract all information about the road-finishing surfaces (RTDwg) along with their shoulders, subassemblies, and ditches. If the road-finishing surfaces are not available, we will apply the rule (created based on engineering knowledge of Komatsu) and approximate the finishing surface. The information and datapoints are save in different csv files; each tagged with a station number
5. Re-creation of the roads centerlines 3D: we will use VB or C#
6. Transform the 3D centerline to alignments: we will use VB.NET or C#.NET
7. Libraries building of the road-transversal-cross-sections 2D: we will use VB or C#.NET
8. Importing EG data: we will use VB.NET or C#.NET
9. Building corridors: we will use VB.NET or C#.NET
10. Creating 3D surface we will use VB.NET or C#.NET        
