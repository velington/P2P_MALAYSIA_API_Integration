# P2P_MALAYSIA_API_Integration
To integrate data from the Courier: Parcel to Post Malaysia into Flat files for SAP Consumption
The program uses Powershell, hence it has no licensing costs.
The key input for the file is a list of consignment nos as csv file. These are the # which the program will fetch sequentially.

The program performs the below steps
1. Logging of the program
2. PLaceholders for final merged csv files, individual csv files, json files, images, logs, data not fetched csv, etc
3. executes API call to P2P server
4. converts json to csv
5. cleans the file of EOF characters, spaces, special characters
6. Downloads image associated with the delivery
7. Loops through the list of all Consignment# and saves each delivery result in separate file
8. Merges all the csv files into 1 merged files with status as completed.
9. Move all the respective files to the respective locations.
