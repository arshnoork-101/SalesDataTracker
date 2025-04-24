# 📊 Sales Report Processing App

A Python-based terminal application that helps you collect, process, and visualize sales data for multiple stores. It generates Excel reports with detailed product-level sales and discount insights, along with easy-to-understand bar charts.

---

## 🚀 Features

- ✅ Interactive terminal input for multiple stores, products, sales, and discounts  
- 📁 Creates individual Excel sheets for each store with:
  - Original sales values  
  - Discounted prices (corrected)  
  - Profit calculation (based on 30% margin)  
- 📊 Automatically generates **bar charts**:
  - One showing original sales  
  - Another showing discounted prices  
- 📈 Summary sheet with:
  - Total sales  
  - Total corrected prices  
  - Total profits across all stores  
- 📤 Clean and organized Excel output using `openpyxl`  



## 🗂️ Project Structure
```
sales_report_project/
│
├── app.py                 # Main application script
├── requirements.txt       # Python dependencies
├── README.md              # Project documentation
└── data/
    ├── store_sales_detailed.xlsx   # Store-wise detailed Excel report
    └── store_sales_summary.xlsx    # Summary report across all stores
```

## 📦 Requirements

All required packages are listed in `requirements.txt`. 
Install them with:

```bash
pip install -r requirements.txt
```


## ▶️ Run the Application

Once the dependencies are installed, follow these steps to run the app:

1. **Open your terminal** or command prompt.

2. **Navigate to the project directory**:

   ```bash
   cd sales_report_project
    ```

3. Run the application
    ```
    python app.py
    ```

4. Follow the On-Screen Prompts to enter the required data for Store Sales

5. After successful execution, two Excel files will be generated inside the `data/` folder:

- `store_sales_detailed.xlsx` — Contains individual sheets for each store with charts and calculated values.
- `store_sales_summary.xlsx` — Contains a summary of total sales, discounted prices, and profits across all stores.
