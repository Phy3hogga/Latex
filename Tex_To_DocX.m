%% Compiles a rudimentary .docx from a .tex file including bibliography and table of contents
% Requires the following programs installed on windows:
% - Texlive / MikTex
% - Pandoc
% Other operating systems are currently unsupported
%% User Input
%Tex File
if(~exist('Tex_File','var') || ~exist('Tex_File_Path','var'))
    [Tex_File, Tex_File_Path] = uigetfile('*.tex','Select Latex .Tex File');
end
%% Get operating system
if(ispc)
    %Call powershell
    Operating_System = "powershell.exe -inputformat none";
elseif(islinux)
    Operating_System = "";
elseif(ismac)
    Operating_System = "";
else
    Operating_System = "";
end
%% Run Command

%Validate tex file
if(isequal(Tex_File, 0))
    clear Tex_File Tex_File_Path;
    error("No Tex File Selected");
else
    %% Bibliography
    if(~exist('Bib_File','var') || ~exist('Bib_File_Path','var'))
        [Bib_File, Bib_File_Path] = uigetfile('*.bib','Select Bibliography .Bib File');
    end
    %Validate if bibliography file is used
    if(isequal(Bib_File, 0))
        Bibliography_Used = false;
        clear Bib_File Bib_File_Path;
    else
        Bibliography_Used = true;
    end
    %Get filename of .tex file without extension
    [~,Tex_File_Name,~] = fileparts(Tex_File);
    %% Build System Command to Pandoc
    %Change directory to tex file directory
    Command = strcat(Operating_System, " cd '", string(Tex_File_Path),"';");
    %Call Pandoc
    Command = strcat(Command, "pandoc -f latex -t docx ", Tex_File, " -o ", Tex_File_Name, ".docx --wrap=preserve --toc");
    %If using a bibliography, add path
    if(Bibliography_Used)
        %Add bibliography
        Command = strcat(Command, " --bibliography '", Bib_File_Path, Bib_File,"'");
        %% Bibliography styling
        %Search for CSL files
        Style_Directory = strcat(fileparts(mfilename('fullpath')), filesep, 'CSL Styling');
        Styling_Contents = dir(strcat(Style_Directory, filesep, '*.csl'));
        %Options to select from
        Option_Names = strrep({Styling_Contents.name},'.csl','');
        %Create figure holding CSL selection
        Selection_Figure = figure('InnerPosition', [0 0 520 380], 'Resize', 'off');
        Control_Search_Label = uicontrol('Style','togglebutton','String','Search:', 'Position', [420 340 80 20]);
        Control_Search_User_Box = uicontrol('Style','edit','String','', 'Position', [20 340 400 20]);
        Search_String_Value = get(Control_Search_User_Box, 'String');
        Search_String_Button = get(Control_Search_Label, 'value');
        Match_Index_List = ~cellfun(@isempty, regexp(Option_Names, strcat('.*?',Search_String_Value,'.*?')));
        Live_Search_Option_Names = Option_Names(Match_Index_List);
        Control_List = uicontrol('Style', 'listbox', 'String', Live_Search_Option_Names, 'Position', [20 40 480 300]);
        Selected_List_Entry = get(Control_List, 'value');
        Control_Button_Select = uicontrol('Style', 'togglebutton', 'String', 'Select', 'Position', [20 20 250 20]);
        Control_Button_Cancel = uicontrol('Style', 'togglebutton', 'String', 'Cancel', 'Position', [250 20 250 20]);
        Button_Select = get(Control_Button_Select, 'value');
        Button_Cancel = get(Control_Button_Cancel, 'value');
        Previous_Search_String = Search_String_Value;
        %While the figure exists
        while ((Button_Select == false) && (Button_Cancel == false))
            %Ensure that the figure can't be closed unexpectedly
            try
                %verify figure exists
                if(~ishandle(Selection_Figure))
                    %Recreate figure if it exists
                    Selection_Figure = figure('InnerPosition', [0 0 520 380], 'Resize', 'off');
                    Control_Search_Label = uicontrol('Style','togglebutton','String','Search:', 'Position', [420 340 80 20]);
                    Control_Search_User_Box = uicontrol('Style','edit','String','', 'Position', [20 340 400 20]);
                    Search_String_Value = get(Control_Search_User_Box, 'String');
                    Search_String_Button = get(Control_Search_Label, 'value');
                    Match_Index_List = ~cellfun(@isempty, regexp(Option_Names, strcat('.*?',Search_String_Value,'.*?')));
                    Live_Search_Option_Names = Option_Names(Match_Index_List);
                    Control_List = uicontrol('Style', 'listbox', 'String', Live_Search_Option_Names, 'Position', [20 40 480 300]);
                    Selected_List_Entry = get(Control_List, 'value');
                    Control_Button_Select = uicontrol('Style', 'togglebutton', 'String', 'Select', 'Position', [20 20 250 20]);
                    Control_Button_Cancel = uicontrol('Style', 'togglebutton', 'String', 'Cancel', 'Position', [250 20 250 20]);
                    Button_Select = get(Control_Button_Select, 'value');
                    Button_Cancel = get(Control_Button_Cancel, 'value');
                end
                %Get current search string
                Search_String_Value = get(Control_Search_User_Box, 'String');
                if(~strcmp(Search_String_Value, Previous_Search_String))
                    Match_Index_List = ~cellfun(@isempty, regexp(Option_Names, strcat('.*?',Search_String_Value,'.*?')));
                    Live_Search_Option_Names = Option_Names(Match_Index_List);
                    Control_List = uicontrol('Style', 'listbox', 'String', Live_Search_Option_Names, 'Position', [20 40 480 300]);
                    Previous_Search_String = Search_String_Value;
                end
                %Allow time to stop force refreshing
                pause(1);
                %Force update on search button
                Search_String_Button = get(Control_Search_Label, 'value');
                if(Search_String_Button)
                    drawnow;
                    set(Control_Search_Label, 'value', 0);
                end
                %Get selected option
                Selected_List_Entry = get(Control_List, 'value');
                Button_Select = get(Control_Button_Select, 'value');
                Button_Cancel = get(Control_Button_Cancel, 'value');
            catch
                %do nothing; figure will recreate itself
            end
        end
        %Close figure automatically on select or close
        close(Selection_Figure);
        %If select button pressed, add csl styling
        if(Button_Select)
            %Get selected option
            Selected_Option = Live_Search_Option_Names{Selected_List_Entry};
            %Add bibliography style
            Command = strcat(Command, " --csl='", Style_Directory, filesep, Selected_Option, "'.csl");
        end
    end
    %End command
    Command = strcat(Command, ";");
end
%Run Powershell command to compile to docx
system(Command);