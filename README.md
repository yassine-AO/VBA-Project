# **VBA Project: Automating Invoicing and Loan Simulation**✨📝

## **Overview**⭐

This VBA project automates the process of invoicing and provides a loan simulation tool. It is designed to streamline client, product, and invoice management, as well as facilitate precise loan simulations.

### **Features**

1. **Client Management**: Add, delete, and manage client details such as name, address, and contact information.
2. **Product Management**: Add, delete, and manage product details including references, designations, and unit prices.
3. **Invoice Generation**: Create detailed invoices with client information, product details, and automatic calculations for total, discount, VAT, and final amount.
4. **Loan Simulation**: Simulate loan repayments based on the invoice amount, annual interest rate, and duration in months.

## **How It Works**⭐

### **User Interface**

- **Home Interface**: Provides buttons to navigate to different functionalities - Clients, Products, Invoice, and Loan Simulation.
- **Client Interface**: Manage client details with options to add, delete, or clear client information.
- **Product Interface**: Manage product details with options to add, delete, or clear product information.
- **Invoice Interface**: Generate invoices by entering client and product details, with automatic calculations for various financial figures.
- **Loan Simulation Interface**: Perform loan simulations using invoice amounts and specified interest rates and durations.

### **Core VBA Procedures**

- **Adding a Client**: Uses **`ajouterclient_Click`** to add a new client.
- **Deleting a Client**: Uses **`CommandButton2_Click`** in the Clients module to delete client entries.
- **Adding a Product**: Uses **`CommandButton2_Click`** in the Products module to add a new product.
- **Deleting a Product**: Uses **`CommandButton2_Click`** in the Products module to delete product entries.
- **Generating an Invoice**: Uses several procedures to gather client and product information, calculate totals, and format the invoice sheet.
- **Running Loan Simulations**: Uses **`lancer_simulation`** to calculate and display loan repayment schedules.

## **Setup Instructions**⭐

1. **Download the Project**: Clone or download the repository from GitHub.
2. **Open the Excel File**: Open the Excel file containing the VBA project.
3. **Enable Macros**: Make sure macros are enabled to allow the VBA scripts to run.

## **Usage Instructions**⭐

### **Managing Clients**

1. Navigate to the Clients interface using the home screen.
2. Use the provided buttons to add or delete clients.
3. Enter the client's name, address, and other details in the appropriate fields.

### **Managing Products**

1. Navigate to the Products interface using the home screen.
2. Use the provided buttons to add or delete products.
3. Enter the product's reference, designation, and unit price in the appropriate fields.

### **Generating Invoices**

1. Navigate to the Invoice interface using the home screen.
2. Enter the client's number to autofill their details.
3. Enter the product references and quantities.
4. The invoice details including totals and taxes will be calculated automatically.

### **Running Loan Simulations**

1. Navigate to the Loan Simulation interface using the home screen.
2. Enter the annual interest rate and the loan duration in months.
3. The tool will automatically use the total from the latest invoice for the loan amount.
4. Run the simulation to see detailed repayment schedules including monthly payments, interest, and remaining balance.

## **Code Highlights**⭐

- **Encapsulation**: Uses **`Private Sub`** to limit the scope of procedures within their respective modules.
- **Variable Declarations**: Utilizes **`Dim`** to declare variables for storing data and manipulating worksheet contents.
- **Worksheet Interaction**: Employs **`Set`** to assign worksheets and ranges for efficient data handling.
- **Loop Structures**: Implements loops like **`For Each`** to iterate over ranges and perform operations on multiple cells.
- **Data Manipulation**: Uses VBA functions to calculate totals, discounts, VAT, and more for accurate financial reporting.

## **License**📜

This project is licensed under the MIT License.
