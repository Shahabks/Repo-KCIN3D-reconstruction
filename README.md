# Repo-KCIN3D-reconstruction
## For the prototype KMCIN1 and feasibility study Feb-March 2021

        On Apr 14 2021
         one File is revised DX_Komatsu_Prototype_0-233D1CSZ_横断図25.ipynb
        
        
        On Apr 7 2021
         One file is revised DX_Komatsu_Prototype_0-183D1PLZ_平面図04_REV1
         
         Added 
         
         >>> DX_Komatsu_Prototype_0-233D1CSZ_横断図25.ipynb
         >>> DFX_DWG2CSV.cs
         >>> vs_community__2111406413.1569763062.exe


*the image shows other notebooks that we need to integrate into the main notebook.*

*Also we need to add markdown boxes to describe the logic + theories + algorithms behind each code-cell. It will make our approach traceable and useful for knowledge-transfer task.*

*.......So please be patient for a couple of days until I get the whole job done .......*

![Image](https://github.com/Shahabks/Repo-KCIN3D-reconstruction/blob/main/Picture1.png)


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
