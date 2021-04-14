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


Approach-1 includes eight steps of which the first four steps are identical to the steps of Approach-2. In this approach, we will try to minimize dependency to C3D functionalities. The outputs of this approach will be the 3D view of the roads-surface (TIN), the 3D roads information in .XML files. The metric that shows the performance of our model/approach will be the distance between the already-available XML data of the roads (as the ground truth) and our model’s generated data outputs. 


         
