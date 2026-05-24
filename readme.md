# About 

The ExSMG is a Excel Add-in assist the business to fastly build the management application based on the advances of Microsoft Excel. It's use javascript to handle the worksheet data, process strings, and interact with HTTP through various built-in functions. Additionally, it offers several utility functions to simplify tasks. While an ERP system is large and complex, making it challenging to modify and adapt for smaller tasks, the ExSMG is a Swiss knife, more adaptable and easier to adjust for smaller tasks based on business demand.

The Ideas of ExSMG Add-in are:

- The business needs to build small or medium-scale tasks instead of developing the application system with a large IDE framework, which has a complicated process and increased costs. This add-in will supply all toolkits necessary for transforming Excel into a simple management application, including both back-end and front-end functionalities if needed.
    
    - The Intellisense box, Worksheet selection box, Filter box, and Command box from the Add-in combine with the Excel data grid structure, including named ranges and table ranges, making Excel a front-end application.
    - Excel has data combined with functions calculating in each cell on the grid table, making Excel an efficient visual database.
    - The Schedule and Command features from the Add-in allow Excel to process complex logic and enable the use of external DLL libraries, extending Excel's interaction capabilities similar to a back-end application.
   
- In ExSMG, Tasks are functions defined by JavaScript. They process data asynchronously, then push the result back to Excel and don't interrupt any Excel operations. The task has 3 types: 
    - Command 
    - Intellisense 
    - Schedule 
- Why is Javascript?. Javascript is simple to learn, good in text processing and JSON is native.
  Javascript is hosted by Javascript V8Engine and intergrated in Microsoft Clearscripts library. JavaScript communicates with Excel through JSON, with the ExSMG serving as the translator between the two. It encodes Excel data into JSON and transmit to the Javascript function and vice versa, provide the wrap functions to effectively handle excel operations.
- Excel-DNA is a .NET Framework library that wraps all Excel C APIs, translating .NET functions into corresponding Excel actions. The ExSMG utilizes Excel-DNA to interact with all Excel components.

Currently, the project has only guide files and plug-in release files. The source code will be fully available soon.

# Note
- The ExSMG add-in mostly work well when the data areas are clearly defined, including named ranges and tables.
- All add-in setting for workbook must be put in a table called ***"config"***.
- Symbols explanation:

    > **[ Button ]** : button and tab button control.
    >
    > **/ Windows Title /** : The windows form with name/title was showed on desktop.
    >
    > **{ Group Name }** : group panel control
    >
    > **Label [ text ]** : Text in textbox control.
    >
    > **Label [ item1, item2 ]** : Items were selected in combobox, listbox control
    >
    > **[ x ] Label** : Radio / Checkbox control
    >


# Installation
1. Download the Add-in released files
2. Extract .zip file to specific directory.
3. In Microsoft Excel
    - [File] > [Options]
    - / Excel Options /
        - [Add-ins]
        - Manage [Excel Add-ins] 
        - [Go...]
            - / Add-ins /
                - [Browse ...]
                - / Browse /
                    - Go to extracted directory above, and choice the ExSMG-Addin64.xll
                    - [OK]
                - { Add-in Available }
                    - [ x ] ExSMG Add-In
                - [OK]

# Features
## <div align="center">COMMAND</div>

### <div align='right'>Default Settings</div>

> *Create the default setting table is called **config***

- On ribbon bar: [ExSMG] > { COMMAND } > [Default Setting]
- The content will be generate as below:

    | config_name | config_value |
    |---|---|
    | tiny_file_path | [ {tiny_fname: "", file_path:"", ole_file:false}, ] |
    | intellisense_map | [ {address:"", tiny_fname:"", function_name:"", trigger_text:"", multi_selection:false}, ] |
    | command_list | [ {description:"", tiny_fname:"", sheet_name:"", function_name:"", include_formula:false} ] |
    | storage_path | .\ |
    | extend_host_cfg | {namespace: "", host_types:[], assemblies: []} |
    | schedule | [ {tiny_fname:"", sheet_name:"", function_name:"", repeat_every:5000} ] |
    | debug_ctp_position | |



### <div align="right">Worksheet Selection</div>
> *Display a dialog for worksheet selection.*

- On ribbon bar: **[ExSMG] > { COMMAND } > [Sheet Selection]**

- Shortcut keys: **Ctrl+Tab**

### <div align="right">Command Dialog</div>
    
> *Display a dialog for selecting commands for execution as configured in the **"config"** table. Once the user selects a command description in the dialog, the add-in will locate and execute the corresponding callback function defined in the JavaScript file.*  

- On ribbon bar: **[ExSMG] > { COMMAND } > [Command Dialog]**

- Shortcut keys: **Ctrl+Shift+C**
    
- Setting in ***"config"*** table as:
    
    - **Config_Name** : command_list
    - **Config_Value** :

        ```json 
        [ 
            {
                description: "Create sale order", // The description that display in Command Dialog.
                tiny_fname:"pizza_sales", // the tiny_fname was defined at  "tiny_file_path" section in "config" table
                function_name: "cmd_CreateSaleOrder", //the call back function is defined in the .js file
                sheet_name:"POS", // Name of worksheet that function work in,
                include_formula:false
            }
        ]
        ```
        

- The callback function in **pizza_sales.js** file: 
    ```javascript
    /**
    * Function description
    * @param   data  worksheet data description was encoded to json
    * @return  null / json - worksheet data encoded for add-in processing
    */
    function cmd_CreateSaleOrder(data) {
    // content ...
    }
    ```

## <div align="center">SQL EXECUTION</div>

This function allow directly execute the SQL command:
- Query records from any data sources, including the self of Ms.Excel file, and view them in an Excel table.
- Execute non-query commands in the data source.
- Transfer records queried from the data source to a specific database destination.

***Attentions:***
- *Currently, this function only supports database drivers including PostgreSQL, SQLite and Oledb driver for ODBC data sources.*
- *When one thread is executing, any attempts to start another thread will be ignored. This thread doesn't interrupt any excel operation.*
- *SQL commands were listed in table will be executed sequentially row by row.*

### <div align="right">SQL Default Setting</div>
> *Create a table that includes the complete default configurations for SQL execution tasks.* 

1. Active any cell in the worksheet.
        
2. On the ribbon bar: **[ExSMG] > {SQL EXECUTION} > [SQL Default Setting]**

    > *The table with default setting is created. If the table was created before, the default setting would be re-filled in the cell at **"setting"** column* 

    The table has 9 fields:
    - **sql_command** : query / non-query to access source table.
    - **params_address** : the address string of range contain pair of params 
    - **target_address** : the address string of table for review data from sql_command above.
    - **setting** : 
        ```json
        {
            non_query: false, 
            delete_all_before: false, // remove all table
            clear_rows_before: false, // remove all rows 
            auto_generate_column: false, // automatically detects new columns from the source and creates corresponding columns in excel table.
            row_by_row_insert: false, // Set to false to optimize the insertion of large volumes of data.
            import: false, // Set true if you want to transfer data from a data source to another data source
            import_table: "", // The table name of destination table
            header_cells: [], // All destination data fields name map with the fields after apply etl from source.
            header_types: [], // Data type of all fields were list in header_cells above
        }
        ```
    - **driver_type** : the driver type of connection string connect to source database
    - **connection_string** : the connection string connect to source database
    - **driver_type_recv** : driver type of connection string connect to destination database
    - **connection_string_recv** : the connection string connect to destination database
    - **work_progress** : state of this progress




### <div align="right">Execute SQL</div>
> *Create a thread to execute the sql command.*
- [Ribbon] > [ExSMG] > {SQL EXECUTION} > [SQL Execute]

### <div align="right">STOP Execution</div>
> *Immediately stop the sql execution thread.*
- [Ribbon] > [ExSMG] > {SQL EXECUTION} > [Stop Execution]

## <div align="center">DEBUG</div>

### <div align="right">SQL Default Setting</div>
>*Display the debug panel for error monitoring.* 

- On ribbon bar: [ExSMG] > {DEBUG} > [Show Debug Panel]

## <div align="center">WORKSHEET FILES</div>
### <div align="right">Show Worksheet Files</div>
>*Display the worksheet file manager.* 
- On ribbon bar: [ExSMG] > {WORKSHEET FILES} > [Show Worksheet Files]

### <div align="right">Create Worksheet GUID</div>
> *Create GUID on current active worksheet.* 
- On ribbon bar: [ExSMG] > {WORKSHEET FILES} > [Create Worksheet GUID]


# Work on
- Microsoft Windows 10 or above
- Microsoft Office 2010 or above

# Dependent on
- .Net Framework 4.8
- Microsoft Clearscripts
- Office Extractor
- Excel-DNA project

