# RPA-CAPSTONE-PROJECT

### **Capstone Project: Calculating BMI Using Web Automation**

**Aim:**  
To automate the process of calculating Body Mass Index (BMI) using a web-based BMI calculator. The automation will extract height and weight data from an Excel file, input these values into the web calculator, retrieve the BMI results, and store them back into the Excel file.



### **Equipment Required:**  
1. UiPath Studio  
2. An Excel file containing Height and Weight data  
3. A web-based BMI calculator (e.g., https://www.calculator.net/bmi-calculator.html)  



### **Installed UiPath Packages:**  
1. **UiPath.Excel.Activities**  
2. **UiPath.WebAPI.Activities**  
3. **UiPath.UIAutomation.Activities**  



### **Procedure:**  

#### **1. Install UiPath and Required Packages:**  
- Download and install UiPath Studio.  
- Create a new project and install the necessary packages (Excel, Web Automation).  



#### **2. Read Data from Excel:**  
- Use the **Excel Application Scope** activity to open the Excel file containing height and weight data.  
- Use the **Read Range** activity to load the data into a DataTable.  



#### **3. Automate Web Interaction to Calculate BMI:**  
- Open the browser using the **Open Browser** activity and navigate to the BMI calculator website.  
- Use the **For Each Row** activity to loop through each record in the DataTable:  
  - Input Height and Weight using **Type Into** activities.  
  - Click the **Calculate** button using the **Click** activity.  
  - Extract the BMI result from the web page using the **Get Text** activity.  



#### **4. Write Data Back to Excel:**  
- Append the extracted BMI value to the corresponding row in the DataTable.  
- Use the **Write Range** activity to save the updated data back to the Excel file.  



#### **5. Run and Test the Workflow:**  
- Execute the workflow to ensure the automation correctly calculates BMI for all rows in the Excel file and updates the results.  



### **UiPath Workflow Description:**  
The UiPath workflow consists of the following main steps:  
1. **Read Input Data:** Read the height and weight data from Excel.  
2. **Open Web Browser:** Open the BMI calculator website.  
3. **Data Entry:** For each record in the DataTable, input the height and weight.  
4. **Calculate BMI:** Extract the BMI result from the webpage.  
5. **Save Results:** Write the BMI values back into the Excel file.

### WorkFlow:
![image](https://github.com/user-attachments/assets/d1429c13-f83a-42d1-8ba0-506685aaeb83)

![image](https://github.com/user-attachments/assets/f849c112-52aa-43e7-aac8-c0def379041f)

![image](https://github.com/user-attachments/assets/a1c2d899-d566-465e-b834-4650491ce946)

![image](https://github.com/user-attachments/assets/20f40f5d-c65a-462a-bf46-ca54ff5fbc9a)

![image](https://github.com/user-attachments/assets/d03c7dea-385f-44e4-9ba9-438ff98988c6)

![image](https://github.com/user-attachments/assets/8844baf9-711d-4146-bd91-630404ab88e3)

![image](https://github.com/user-attachments/assets/a0c428b7-af7f-4308-ba5b-5aefb4790c72)

![image](https://github.com/user-attachments/assets/08db09b6-763f-47c9-a044-f1f2bee0b546)

### OUTPUT

![Screenshot 2024-11-24 162044](https://github.com/user-attachments/assets/114ea80b-d187-468d-94c2-77fa6abad1bc)

![Screenshot 2024-11-24 162102](https://github.com/user-attachments/assets/ba1fb932-34a1-4785-8d4c-7e3274bd86d0)

![Screenshot 2024-11-24 162118](https://github.com/user-attachments/assets/0dd94d57-d194-41d0-8157-9aa1cf5edffe)

![Screenshot 2024-11-24 162131](https://github.com/user-attachments/assets/afb7e82d-9835-4f70-a7b2-58a66b918b15)

![Screenshot 2024-11-24 162146](https://github.com/user-attachments/assets/a061b94d-df51-4e16-b965-144d28afd187)

### **Result:**  
The automation successfully reads the height and weight data from Excel, calculates BMI using a web calculator, and stores the BMI results in the Excel file. This demonstrates the practical use of web automation for health-related calculations.  



