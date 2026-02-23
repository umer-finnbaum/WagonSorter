# WagonSorter

This is a trimmed version of the Wagon Sorter with no identifier data. Comments are also kep to a minimum.  

Please note that some functionality might be broken in the process of trimming for public release.  

- This program uses QtWidgets due to performance requirements. Future versions will exclusively use Tkinter.
- Data file is created in CX-supervisor or any other SCADA platform. This program uses configuration of wagons to create a sorting logic for each part.
- Each part has a unique identifier, a "VIP-key" from a predefined set, and dimensions.  
- Final output provides Wagon number, row, and place on row, for each part as a text file.
- This repo includes the simulator script for testing without a PLC.
- This is the final version program currently in production. There is redundant code left over from previous iterations. The code has evolved over several months of requirement changes from the end users.  
- Included txt files are examples. These files need to be placed in "C:/FBData/Stacker/" or change the location in relevant script.  



## Stacker

Takes input from settings.txt for Wagon dimensions:  
`wagon_number, rows, slots_per_row, total_width, total_height, allowed_vip_keys, allow_vip_key_mixing, stacking_per_slot`  
-  "allow_vip_key_mixing" is set to either yes or no. "yes" allows different VIP-keys to be placed on the same row.  
-  "stacking_per_slot" is set to either 1 or 2. Setting it to 2 allows the parts to be stacked vertically on the slot.

Takes input from filename_final.txt for incoming parts, writes back to the same file:  
`Placed,Wagon,Row,Slot,SerieID,Serie,VIPKEY,TOOLPACK1,TOOLPROF1,TOOLPACK2,TOOLPROF2,MaterialType,MouldedWidth,MouldedHeight,ProfType,PartLength,PartNum,Pol,PairNum,FinishedLength`  
-  This data file is necessary for sorting logic.
-  filename is configurable. Polled from PLC at a specified STRING[256] tag.
-  Program writes to this file atomically. It is possible to resume from any point.
-  filename_final.txt is generated via wagon_sorter script in Supervisor and the file needs to be copied to PC running AutoStacker.
-  Previously sorting was done completely by AutoStacker, this change ensures a fresh file is copied on new series start.


**Main Window with part details (Single Wagon display):**  

Single Wagon display (Updated to horizontal text)  

<img width="1918" height="1026" alt="W1" src="https://github.com/user-attachments/assets/5006218e-b0a4-4ffa-8bdd-06ca1a08d664" />


Dual Wagon display (Updated to horizontal text)  

<img width="1692" height="896" alt="W2" src="https://github.com/user-attachments/assets/1d27595a-d31d-4ad5-88a8-ea638a8e2e8f" />


------------------------------

**Overview Window with all wagons:**  

Changed in latest update to show only 2 wagons at a time.  

<img width="2559" height="1389" alt="image" src="https://github.com/user-attachments/assets/a95c10bd-1f6a-43fb-be94-e800990a29b3" />


## Remover (WIP)
(Currently no plans to integrate this to main Stacker visualizer.)  
Manual control only. Can be modified to be automated based on requirements.  

Takes input from output file of Stacker (Data_out.txt).  
Writes output to shapes_remaining.txt if the **Save Remaining** button is pressed. Format is preserved.  

## Temporary Helper function for PLC Read/Write
Basic read/write function for debugging. Uses aphyt for writing string to Omron controllers.  

![PLCRW](https://github.com/user-attachments/assets/f053c9d9-500f-4fa6-8fa5-a266f43618c7)

