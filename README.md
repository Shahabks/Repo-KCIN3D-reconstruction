# Repo-KCIN3D-reconstruction
## For the prototype KMCIN1 and feasibility study Feb-March 2021
        On Apr 16 2021
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


*the image shows other notebooks that we need to integrate into the main notebook.*

*Also we need to add markdown boxes to describe the logic + theories + algorithms behind each code-cell. It will make our approach traceable and useful for knowledge-transfer task.*

*.......So please be patient for a couple of days until I get the whole job done .......*

![Image](https://github.com/Shahabks/Repo-KCIN3D-reconstruction/blob/main/Picture1.png)

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
