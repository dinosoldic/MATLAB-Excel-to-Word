clear;clc;
% Read student information from Excel using readtable
[ExcelFile, ExcelFilePath] = uigetfile({'*.xlsx', 'MS Excel xlsx';'*.xls', 'MS Excel xls'},"Select Excel Database");
ExcelData = readtable(fullfile(ExcelFilePath, ExcelFile)); 

% Define the Word template file and output directory
[templateFile, templateFilePath] = uigetfile({'*.docx', 'MS Word docx';'*.doc', 'MS Word doc'},"Select a Word Template"); 

% Set save options
prompt = {'Enter Save Folder Name:', 'Enter Save File Name:'};
dlgtitle = "Save options:";
dims = [1 60]; 
saveFolder = inputdlg(prompt, dlgtitle, dims);

% Prompt to save as pdf
saveAsPDF = string(questdlg("Do you wish to also save as '.pdf'", "Save as PDF", "Yes", "No", "Yes"));

% Prompt user to select save folder
msg = "Next, select a folder to save your formatted data.";
uiwait(warndlg(msg, 'Folder Selection'));
outputDir = uigetdir("Select Folder to Save Data");
outputNewDir = fullfile(outputDir, saveFolder{1});
saveFullFile = fullfile(outputNewDir, saveFolder{2}); 

% Create output directory if it doesn't exist
if ~exist(outputNewDir, "dir")
    mkdir(outputNewDir);
end

% Set placeholders in Word
placeholder = struct();
currentPlaceholderOrder = cell(1, numel(ExcelData(1, :)));
placeholderIndex = 1;

while true
    prompt = "Enter Expression to Replace in Word:";
    dlgtitle = "Add Placeholders:";
    dims = [1 50];
    addPlaceholder = inputdlg(prompt, dlgtitle, dims);

    % Update Placeholders
    placeholder(placeholderIndex).expression = string(addPlaceholder{1});
    currentPlaceholderOrder{placeholderIndex} = addPlaceholder{1}; 
    placeholderIndex = placeholderIndex + 1;

    % Prompt User to add more and wanr user
    continueAddPlaceholder = string(questdlg("Do you wish to replace any other expression?", "Add More Expressions", "Yes", "No", "No"));
    if strcmpi(continueAddPlaceholder, "No")

        currentPlaceholderOrderString = strjoin(currentPlaceholderOrder, ', ');
        warnOrderPlaceholderMsg = "WARNING: Make sure that the expressions order is the same as in the Excel File." + newline +...
            "Your Current order is {" + currentPlaceholderOrderString + "}" + newline +...
            "                                                      " +...
            "Redo Order?" + newline;
        warnOrderPlaceholder = string(questdlg(warnOrderPlaceholderMsg, "Order Warning", "Yes", "No", "No"));

        if strcmpi(warnOrderPlaceholder, "No")
            break
        elseif strcmpi(warnOrderPlaceholder, "Yes")
            warnOrderMsg = "You have to start over." + newline + "Make sure to follow the correct order.";
            uiwait(warndlg(warnOrderMsg, 'Order Warning'));

            placeholder = struct();
            currentPlaceholderOrder = cell(1, numel(ExcelData(1, :)));
            placeholderIndex = 1;
        end
    end
end

% Determine all excel fields
excelFieldNames = struct();
for excelMaxWidth_Col = 1:width(ExcelData)
    excelFieldNames(excelMaxWidth_Col).fieldnames = char(ExcelData.Properties.VariableNames{excelMaxWidth_Col});
end

% Convert excel table to atring arrays to ease looping
ExcelData = table2cell(ExcelData);

% Access Word app
wordApp = actxserver("Word.Application");
wordApp.Visible = 1;

for excelLoop_Row = 1:height(ExcelData) % Determine n rows
    % Open the Word template
    doc = wordApp.Documents.Open(fullfile(templateFilePath, templateFile));
    
    % Loop through all fields and replace placeholders with excel data
    for excelMaxWidth_Col = 1:width(ExcelData)
        content = doc.Content;

        while content.Find.Execute(placeholder(excelMaxWidth_Col).expression)
            content.Find.Execute(placeholder(excelMaxWidth_Col).expression)
            if  isnumeric(ExcelData{excelLoop_Row, excelMaxWidth_Col})
                content.Text = num2str(ExcelData{excelLoop_Row, excelMaxWidth_Col});                
            else
                content.Text = string(ExcelData{excelLoop_Row, excelMaxWidth_Col});                
            end           
            content = doc.Content;
        end
    end

    % Save doc with custom name
    customSaveAs = string(saveFullFile) + "_" +string(ExcelData{excelLoop_Row, 1});
    doc.SaveAs2(customSaveAs);

    % Additionally save as pdf
    if strcmpi(saveAsPDF, "Yes")
        doc.SaveAs2(customSaveAs, 17);
    end
    
    % Close the document without saving changes to the template
    doc.Close(0);
end

% Close Word application
wordApp.Quit;

% Display completion 
totfiles = excelLoop_Row;
completionMsg = sprintf("\n\n-------Successfully completed %d files-------\n", totfiles);
catASCII = "/\\_/\\       /\\_/\\" + newline +"( o.o )     ( ^.^ )" + newline + "> ^ <       > ^ <";

% Calculate screen size
screenSize = get(0, 'ScreenSize');
screenWidth = screenSize(3);
screenHeight = screenSize(4);

% Create a figure for the completion message
fig = figure('Name', 'Process Completed', 'NumberTitle', 'off', 'MenuBar', 'none');
textHandle = uicontrol('Style', 'text', 'String', [completionMsg, catASCII], ...
    'HorizontalAlignment', 'center', 'FontSize', 12, 'FontWeight', 'bold', ...
    'Units', 'pixels');

% Calculate figure position to center on the screen
figureWidth = 400; % Adjust figure width
figureHeight = 200; % Adjust figure height
figureX = (screenWidth - figureWidth) / 2;
figureY = (screenHeight - figureHeight) / 2;

% Set figure position
set(fig, 'Position', [figureX, figureY, figureWidth, figureHeight]);

% Set text position within the figure
set(textHandle, 'Position', [0, 0, figureWidth, figureHeight]);

