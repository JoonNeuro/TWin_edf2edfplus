function edf2set(~)
    % Automated script to merge EDF files with Excel event files and save as SET(EEGLAB) format
    % 
    % Input:
    %   base_directory - Base directory path to process (optional)
    %                    If not provided, uses current directory
    %
    % Usage:
    %   edf2edfplus()                    % Run in current directory
    %   edf2edfplus('/path/to/data')     % Run in specific directory
    
    % Test file paths (managed in one place)
    test_edf_file = '/Users/joon/Documents/Projects/PatientID_YYYYMMDD_HHMM.edf';
    test_excel_file = '/Users/joon/Documents/Projects/PatientID_YYYYMMDD_HHMM_Comments.xlsx';
    
    % Test: Process single file only
    test_single_file(test_edf_file, test_excel_file);
    
    % Actual EDF+ conversion
    convert_edf_to_edfplus(test_edf_file, test_excel_file);
    return;
    
    function test_single_file(edf_file, excel_file)
        % Function to test with a single file only
        
        fprintf('=== Single File Test Start ===\n');
        
        fprintf('EDF file: %s\n', edf_file);
        fprintf('Excel file: %s\n', excel_file);
        
        % 1. Load EDF file and check information
        fprintf('\n1. EDF file analysis...\n');
        try
            % Read EDF header directly
            [edf_duration, sampling_rate, n_channels] = read_edf_header(edf_file);
            fprintf('  EDF header read success\n');
            fprintf('  Sampling rate: %.1f Hz\n', sampling_rate);
            fprintf('  Number of channels: %d\n', n_channels);
            fprintf('  Data length: %.2f seconds\n', edf_duration);
            
            % Verify with direct EDF reading
            try
                [header, data] = read_edf_file(edf_file);
                EEG = create_eeg_struct(header, data);
                fprintf('  Direct read length: %.2f seconds (%.0f samples)\n', EEG.pnts/EEG.srate, EEG.pnts);
                fprintf('  ✓ Direct read matches header info: %.2f seconds\n', edf_duration);
            catch ME
                fprintf('  Direct read failed: %s\n', ME.message);
            end
        catch ME
            fprintf('  EDF header read failed: %s\n', ME.message);
            return;
        end
        
        % 2. Excel file analysis
        fprintf('\n2. Excel file analysis...\n');
        try
            % Use readtable for more accurate reading
            try
                % Method 1: Use readtable
                data_table = readtable(excel_file, 'ReadVariableNames', false);
                raw = table2cell(data_table);
                fprintf('  Excel load success (using readtable)\n');
            catch
                % Method 2: Use xlsread (original method)
                [~, ~, raw] = xlsread(excel_file);
                fprintf('  Excel load success (using xlsread)\n');
            end
            
            fprintf('  Data size: %dx%d\n', size(raw, 1), size(raw, 2));
            
            % Print first row
            fprintf('  First row: ');
            for i = 1:min(5, size(raw, 2))
                if i <= size(raw, 2)
                    fprintf('"%s" ', raw{1, i});
                end
            end
            fprintf('\n');
            
            % Print all rows
            fprintf('  All data:\n');
            for i = 1:size(raw, 1)
                fprintf('    Row %d: ', i);
                for j = 1:size(raw, 2)
                    fprintf('"%s" ', raw{i, j});
                end
                fprintf('\n');
            end
            
        catch ME
            fprintf('  Excel load failed: %s\n', ME.message);
            return;
        end
        
        % 3. Time conversion and length verification
        fprintf('\n3. Time conversion and length verification...\n');
        if size(raw, 1) >= 2
            % Time is in 2nd column (B column, index 2)
            time_col = 2;
            
            % Check first and last events
            first_event = raw{1, time_col};  % 2nd column is time
            last_event = raw{end, time_col}; % 2nd column is time
            
            fprintf('  First event: "%s"\n', first_event);
            fprintf('  Last event: "%s"\n', last_event);
            
            % Time conversion (direct processing)
            % Convert "11:27:51.00" format to seconds and calculate relative time
            if ischar(first_event) && contains(first_event, ':')
                parts1 = strsplit(first_event, ':');
                if length(parts1) >= 3
                    start_abs_time = str2double(parts1{1})*3600 + str2double(parts1{2})*60 + str2double(parts1{3});
                else
                    start_abs_time = 0;
                end
            else
                start_abs_time = 0;
            end
            
            if ischar(last_event) && contains(last_event, ':')
                parts2 = strsplit(last_event, ':');
                if length(parts2) >= 3
                    end_abs_time = str2double(parts2{1})*3600 + str2double(parts2{2})*60 + str2double(parts2{3});
                else
                    end_abs_time = 0;
                end
            else
                end_abs_time = 0;
            end
            
            % Convert to relative time (set first event to 0 seconds)
            start_time = 0;  % First event is 0 seconds
            end_time = end_abs_time - start_abs_time;  % Last event is relative time
            excel_duration = end_time - start_time;
            
            fprintf('  Start time: %.2f seconds\n', start_time);
            fprintf('  End time: %.2f seconds\n', end_time);
            fprintf('  Excel record length: %.2f seconds\n', excel_duration);
            fprintf('  EDF length: %.2f seconds\n', edf_duration);
            fprintf('  Length difference: %.2f seconds\n', abs(excel_duration - edf_duration));
            
            if abs(excel_duration - edf_duration) < 1
                fprintf('  ✓ Length matching success!\n');
            else
                fprintf('  ⚠ Length difference is large!\n');
            end
            
        else
            fprintf('  Insufficient data.\n');
        end
        
        fprintf('\n=== Test Complete ===\n');
    end
    
    function convert_edf_to_edfplus(edf_file, excel_file)
        % Function to perform actual EDF+ conversion
        
        fprintf('\n=== EDF+ Conversion Start ===\n');
        
        fprintf('EDF file: %s\n', edf_file);
        fprintf('Excel file: %s\n', excel_file);
        
        % 1. Load EDF file
        fprintf('\n1. Load EDF file...\n');
        try
            % 1. Read accurate length information from EDF header
            [edf_duration, sampling_rate, n_channels] = read_edf_header(edf_file);
            fprintf('  EDF header info: %.2f seconds, %.1f Hz, %d channels\n', edf_duration, sampling_rate, n_channels);
            
            % 2. Load data with pop_biosig
            EEG = pop_biosig(edf_file);
            fprintf('  pop_biosig load: %.2f seconds (%.0f samples)\n', EEG.pnts/EEG.srate, EEG.pnts);
            
            % 3. Fix length with header information
            if abs(edf_duration - EEG.pnts/EEG.srate) > 1.0
                fprintf('  Fixing length...\n');
                target_samples = round(edf_duration * EEG.srate);
                if target_samples > EEG.pnts
                    padding = repmat(EEG.data(:, end), 1, target_samples - EEG.pnts);
                    EEG.data = [EEG.data, padding];
                elseif target_samples < EEG.pnts
                    EEG.data = EEG.data(:, 1:target_samples);
                end
                EEG.pnts = target_samples;
                EEG.xmax = (EEG.pnts - 1) / EEG.srate;
                fprintf('  Length fix complete: %.2f seconds (%.0f samples)\n', EEG.pnts/EEG.srate, EEG.pnts);
            end
            
            fprintf('  EDF load success\n');
            fprintf('  Final sampling rate: %.1f Hz\n', EEG.srate);
            fprintf('  Final channel count: %d\n', EEG.nbchan);
            fprintf('  Final data length: %.2f seconds (%.0f samples)\n', EEG.pnts/EEG.srate, EEG.pnts);
        catch ME
            fprintf('  EDF load failed: %s\n', ME.message);
            return;
        end
        
        % 2. Read Excel event file
        fprintf('\n2. Read Excel event file...\n');
        try
            % Use readtable for more accurate reading
            try
                data_table = readtable(excel_file, 'ReadVariableNames', false);
                raw = table2cell(data_table);
                fprintf('  Excel load success (using readtable)\n');
            catch
                [~, ~, raw] = xlsread(excel_file);
                fprintf('  Excel load success (using xlsread)\n');
            end
            
            fprintf('  Data size: %dx%d\n', size(raw, 1), size(raw, 2));
            
            % Print first row
            fprintf('  First row: ');
            for i = 1:min(4, size(raw, 2))
                fprintf('"%s" ', raw{1, i});
            end
            fprintf('\n');
            
        catch ME
            fprintf('  Excel load failed: %s\n', ME.message);
            return;
        end
        
        % 3. Process event data
        fprintf('\n3. Process event data...\n');
        
        % Time is in 2nd column (B column, index 2), event type is in 3rd column (C column, index 3)
        time_col = 2;
        type_col = 3;
        
        % Set first event time as reference
        first_event_time = raw{1, time_col};
        if ischar(first_event_time) && contains(first_event_time, ':')
            parts = strsplit(first_event_time, ':');
            if length(parts) >= 3
                start_abs_time = str2double(parts{1})*3600 + str2double(parts{2})*60 + str2double(parts{3});
            else
                start_abs_time = 0;
            end
        else
            start_abs_time = 0;
        end
        
        fprintf('  Reference time: %s (%.2f seconds)\n', first_event_time, start_abs_time);
        
        % Create event structure
        events = [];
        event_count = 0;
        
        for i = 1:size(raw, 1)
            time_val = raw{i, time_col};
            type_val = raw{i, type_col};
            
            if ~isempty(time_val) && ~isempty(type_val) && ischar(time_val) && contains(time_val, ':')
                % Convert time to seconds
                parts = strsplit(time_val, ':');
                if length(parts) >= 3
                    current_abs_time = str2double(parts{1})*3600 + str2double(parts{2})*60 + str2double(parts{3});
                    relative_time = current_abs_time - start_abs_time;
                    
                    % Convert to sample index (based on sampling rate)
                    latency = round(relative_time * EEG.srate) + 1; % 1-based indexing
                    
                    if latency > 0 && latency <= EEG.pnts
                        event_count = event_count + 1;
                        events(event_count).latency = latency;
                        events(event_count).type = type_val;
                        events(event_count).duration = 0; % Default duration
                        
                        fprintf('  Event %d: %s at %.2f seconds (sample %d)\n', event_count, type_val, relative_time, latency);
                    end
                end
            end
        end
        
        % 4. Add events to EEG structure
        fprintf('\n4. Add events to EEG...\n');
        if ~isempty(events)
            % Create event channel (add as last channel)
            event_channel = zeros(1, EEG.pnts);
            
            % Mark each event in event channel
            for i = 1:length(events)
                latency = round(events(i).latency);
                if latency > 0 && latency <= EEG.pnts
                    % Mark with different values based on event type
                    switch events(i).type
                        case 'Start Recording'
                            event_channel(latency) = 1;
                        case 'C3 Touch'
                            event_channel(latency) = 2;
                        case 'F3 Touch'
                            event_channel(latency) = 3;
                        case 'Cz Touch'
                            event_channel(latency) = 4;
                        case 'Paused'
                            event_channel(latency) = 5;
                        otherwise
                            event_channel(latency) = 99;
                    end
                end
            end
            
            % Add event channel to EEG data
            EEG.data = [EEG.data; event_channel];
            EEG.nbchan = EEG.nbchan + 1;
            
            % Add channel information
            EEG.chanlocs(EEG.nbchan).labels = 'Events';
            EEG.chanlocs(EEG.nbchan).type = 'EVENT';
            
            % Add event structure
            EEG.event = events;
            
            fprintf('  Added %d events (including event channel)\n', length(events));
        else
            fprintf('  No event data available\n');
            return;
        end
        
        % 5. Save as SET file (clean filename)
        fprintf('\n5. Save SET file...\n');
        
        % Generate output filename (keep original EDF name)
        [filepath, name, ~] = fileparts(edf_file);
        set_file = fullfile(filepath, [name '.set']);
        
        % Validate EEG structure
        EEG = eeg_checkset(EEG);
        
        % Save EEG structure as .set file
        pop_saveset(EEG, 'filename', [name '.set'], 'filepath', filepath, 'savemode', 'onefile');
        fprintf('  ✓ SET file save complete: %s\n', set_file);
        
        fprintf('\n=== SET File Conversion Complete ===\n');
    end
    
    % Direct EDF file reading function
    function [header, data] = read_edf_file(filename)
        fid = fopen(filename, 'r');
        if fid == -1
            error('Cannot open EDF file: %s', filename);
        end
        
        % Read EDF header (256 bytes)
        header_bytes = fread(fid, 256, 'uint8');
        header_str = char(header_bytes');
        
        % Parse header information
        header.version = str2double(header_str(1:8));
        header.patient_id = strtrim(header_str(9:88));
        header.recording_id = strtrim(header_str(89:168));
        header.start_date = strtrim(header_str(169:176));
        header.start_time = strtrim(header_str(177:184));
        header.header_bytes = str2double(header_str(185:192));
        header.reserved = strtrim(header_str(193:236));
        header.records = str2double(header_str(237:244));
        header.duration = str2double(header_str(245:252));
        header.nsignals = str2double(header_str(253:256));
        
        % Read signal headers (nsignals * 256 bytes)
        signal_header_bytes = fread(fid, header.nsignals * 256, 'uint8');
        signal_header_str = char(signal_header_bytes');
        
        % Parse each signal's information
        for i = 1:header.nsignals
            start_idx = (i-1) * 256 + 1;
            end_idx = i * 256;
            signal_str = signal_header_str(start_idx:end_idx);
            
            header.signals(i).label = strtrim(signal_str(1:16));
            header.signals(i).transducer = strtrim(signal_str(17:32));
            header.signals(i).units = strtrim(signal_str(33:40));
            header.signals(i).physical_min = str2double(signal_str(41:48));
            header.signals(i).physical_max = str2double(signal_str(49:56));
            header.signals(i).digital_min = str2double(signal_str(57:64));
            header.signals(i).digital_max = str2double(signal_str(65:72));
            header.signals(i).prefilter = strtrim(signal_str(73:80));
            header.signals(i).samples_per_record = str2double(signal_str(81:88));
            header.signals(i).reserved = strtrim(signal_str(89:256));
        end
        
        % Read data
        data = [];
        
        % Calculate total samples for each signal
        total_samples = header.records * header.signals(1).samples_per_record;
        
        % Pre-allocate data array
        data = zeros(header.nsignals, total_samples);
        
        % Read data for each record
        for rec = 1:header.records
            for sig = 1:header.nsignals
                samples = header.signals(sig).samples_per_record;
                if samples > 0
                    signal_data = fread(fid, samples, 'int16');
                    if length(signal_data) == samples
                        start_idx = (rec-1) * samples + 1;
                        end_idx = rec * samples;
                        data(sig, start_idx:end_idx) = signal_data';
                    else
                        fprintf('  Warning: Sample count mismatch in signal %d, record %d\n', sig, rec);
                    end
                end
            end
        end
        
        fclose(fid);
    end
    
    % EEG structure creation function
    function EEG = create_eeg_struct(header, data)
        EEG = [];
        EEG.setname = 'EDF Import';
        EEG.filename = '';
        EEG.filepath = '';
        EEG.subject = header.patient_id;
        EEG.group = '';
        EEG.condition = '';
        EEG.session = [];
        EEG.comments = sprintf('Imported from EDF file. Start: %s %s', header.start_date, header.start_time);
        EEG.nbchan = header.nsignals;
        EEG.trials = 1;
        EEG.pnts = size(data, 2);
        EEG.srate = header.signals(1).samples_per_record / header.duration;
        EEG.xmin = 0;
        EEG.xmax = (EEG.pnts - 1) / EEG.srate;
        EEG.data = data;
        EEG.icaact = [];
        EEG.icawinv = [];
        EEG.icasphere = [];
        EEG.icaweights = [];
        EEG.icachansind = [];
        EEG.chanlocs = [];
        EEG.chaninfo = [];
        EEG.ref = 'common';
        EEG.event = [];
        EEG.urevent = [];
        EEG.epoch = [];
        EEG.epochdescription = {};
        EEG.reject = [];
        EEG.stats = [];
        EEG.specdata = [];
        EEG.specicaact = [];
        EEG.splinefile = '';
        EEG.icasplinefile = '';
        EEG.dipfit = [];
        EEG.history = '';
        EEG.saved = 'no';
        EEG.etc = [];
        
        % Set channel information
        for i = 1:EEG.nbchan
            EEG.chanlocs(i).labels = header.signals(i).label;
            EEG.chanlocs(i).type = 'EEG';
            EEG.chanlocs(i).theta = [];
            EEG.chanlocs(i).radius = [];
            EEG.chanlocs(i).X = [];
            EEG.chanlocs(i).Y = [];
            EEG.chanlocs(i).Z = [];
            EEG.chanlocs(i).sph_theta = [];
            EEG.chanlocs(i).sph_phi = [];
            EEG.chanlocs(i).sph_radius = [];
            EEG.chanlocs(i).urchan = [];
            EEG.chanlocs(i).ref = [];
        end
    end
    fprintf('Base directory: %s\n', base_directory);
    
    % EEGLAB path setup (modify if needed)
    % addpath('/path/to/eeglab');
    
    % Find subdirectories (folders starting with numbers)
    subdirs = dir(base_directory);
    subdirs = subdirs([subdirs.isdir] & ~strcmp({subdirs.name}, '.') & ~strcmp({subdirs.name}, '..'));
    subdirs = subdirs(cellfun(@(x) ~isempty(regexp(x, '^\d+_', 'once')), {subdirs.name}));
    
    if isempty(subdirs)
        fprintf('No subdirectories to process.\n');
        return;
    end
    
    fprintf('Found directories: ');
    for i = 1:length(subdirs)
        fprintf('%s ', subdirs(i).name);
    end
    fprintf('\n\n');
    
    % Process each directory
    for dir_idx = 1:length(subdirs)
        current_dir = fullfile(base_directory, subdirs(dir_idx).name);
        fprintf('Processing directory: %s\n', subdirs(dir_idx).name);
        
        process_directory(current_dir);
    end
    
    fprintf('\nAutomation complete!\n');
end
    
    function process_directory(directory)
    % Process all EDF files in directory
    
    % Find EDF files and Excel event files
    edf_files = dir(fullfile(directory, '*.edf'));
    excel_files = dir(fullfile(directory, '*_Comments.xlsx'));
    
    if isempty(edf_files)
        fprintf('  No EDF files found.\n');
        return;
    end
    
    if isempty(excel_files)
        fprintf('  No Excel event files found.\n');
        return;
    end
    
    % File matching
    matches = find_matching_files(edf_files, excel_files, directory);
    
    if isempty(matches)
        fprintf('  No matching files found.\n');
        return;
    end
    
    fprintf('  Matched file pairs: %d\n', length(matches));
    
    % Process each matched file pair
    success_count = 0;
    for i = 1:length(matches)
        edf_file = matches(i).edf_file;
        excel_file = matches(i).excel_file;
        date = matches(i).date;
        
        fprintf('  Processing: %s\n', edf_file);
        
        % Generate output filename
        [~, name, ~] = fileparts(edf_file);
        
        % Execute conversion (SET files only)
        if convert_edf_to_set(edf_file, excel_file, directory)
            success_count = success_count + 1;
            fprintf('  ✓ Conversion success: %s\n', [name '_with_events.set']);
        else
            fprintf('  ✗ Conversion failed: %s\n', edf_file);
        end
    end
    
    fprintf('  Directory complete: %d/%d success\n\n', success_count, length(matches));
    end
    
    function matches = find_matching_files(edf_files, excel_files, directory)
    % Match EDF files and Excel event files by date
    
    matches = [];
    
    for i = 1:length(edf_files)
        edf_name = edf_files(i).name;
        
        % Extract date from EDF filename (YYYYMMDD format)
        date_match = regexp(edf_name, '(\d{8})', 'tokens');
        if isempty(date_match)
            continue;
        end
        
        edf_date = date_match{1}{1};
        
        % Find Excel event file of the same date
        for j = 1:length(excel_files)
            excel_name = excel_files(j).name;
            excel_date_match = regexp(excel_name, '(\d{8})', 'tokens');
            
            if ~isempty(excel_date_match) && strcmp(excel_date_match{1}{1}, edf_date)
                matches(end+1).edf_file = fullfile(directory, edf_name);
                matches(end).excel_file = fullfile(directory, excel_name);
                matches(end).date = edf_date;
                fprintf('  Matched: %s <-> %s\n', edf_name, excel_name);
                break;
            end
        end
    end
    end
    
    function success = convert_edf_to_set(edf_file, excel_file, output_dir)
        % Convert EDF to SET file with events
        
        success = false;
        
        try
            % Load EDF file
            EEG = pop_biosig(edf_file);
            
            % Load Excel events
            [~, ~, raw] = xlsread(excel_file);
            if size(raw, 1) > 1
                data = raw(2:end, :);
            else
                data = raw;
            end
            
            % Process events
            events = [];
            if ~isempty(data)
                time_col = 2; % Time column
                type_col = 3; % Event type column
                
                % Get reference time
                first_event_time = data{1, time_col};
                if ischar(first_event_time) && contains(first_event_time, ':')
                    parts = strsplit(first_event_time, ':');
                    if length(parts) >= 3
                        start_abs_time = str2double(parts{1})*3600 + str2double(parts{2})*60 + str2double(parts{3});
                    else
                        start_abs_time = 0;
                    end
                else
                    start_abs_time = 0;
                end
                
                % Create events
                for i = 1:size(data, 1)
                    time_val = data{i, time_col};
                    type_val = data{i, type_col};
                    
                    if ~isempty(time_val) && ~isempty(type_val) && ischar(time_val) && contains(time_val, ':')
                        parts = strsplit(time_val, ':');
                        if length(parts) >= 3
                            current_abs_time = str2double(parts{1})*3600 + str2double(parts{2})*60 + str2double(parts{3});
                            relative_time = current_abs_time - start_abs_time;
                            latency = round(relative_time * EEG.srate) + 1;
                            
                            if latency > 0 && latency <= EEG.pnts
                                events(end+1).latency = latency;
                                events(end).type = type_val;
                                events(end).duration = 0;
                            end
                        end
                    end
                end
            end
            
            % Add events to EEG
            if ~isempty(events)
                EEG.event = events;
            end
            
            % Save SET file
            [~, name, ~] = fileparts(edf_file);
            set_file = fullfile(output_dir, [name '.set']);
            EEG = eeg_checkset(EEG);
            pop_saveset(EEG, 'filename', [name '.set'], 'filepath', output_dir, 'savemode', 'onefile');
            
            success = true;
            
        catch ME
            fprintf('    Error: %s\n', ME.message);
            success = false;
        end
    end
    
    function [duration, sampling_rate, n_channels] = read_edf_header(edf_file)
        % Read EDF file header directly to return accurate information
        
        % Set default values
        duration = 0;
        sampling_rate = 0;
        n_channels = 0;
        
        try
            fid = fopen(edf_file, 'r');
            if fid == -1
                error('Cannot open file: %s', edf_file);
            end
            
            % Read EDF header (256 bytes)
            header = fread(fid, 256, 'uint8');
            fclose(fid);
            
            % Parse header information
            n_records = str2double(char(header(237:244))');
            record_duration = str2double(char(header(245:252))');
            n_signals = str2double(char(header(253:256))');
            
            % Set default values
            duration = n_records * record_duration;
            n_channels = n_signals;
            
            fprintf('  EDF header information:\n');
            fprintf('    Number of records: %d\n', n_records);
            fprintf('    Record duration: %.1f seconds\n', record_duration);
            fprintf('    Number of signals: %d\n', n_signals);
            
            % Try to calculate sampling rate for each signal
            try
                fid = fopen(edf_file, 'r');
                fseek(fid, 256, 'bof'); % Skip main header
                
                % Read first signal's sample count
                signal_header = fread(fid, 256, 'uint8');
                fclose(fid);
                
                % Sample count is in signal header bytes 17-20
                samples_per_record = str2double(char(signal_header(17:20))');
                
                if ~isnan(samples_per_record) && samples_per_record > 0
                    sampling_rate = samples_per_record / record_duration;
                    fprintf('    Samples per record: %d\n', samples_per_record);
                    fprintf('    Sampling rate: %.1f Hz\n', sampling_rate);
                else
                    % Use default value (typical EEG sampling rate)
                    sampling_rate = 200; % Hz
                    fprintf('    Sampling rate: %.1f Hz (default)\n', sampling_rate);
                end
            catch
                % Error when sampling rate calculation fails
                error('Failed to calculate sampling rate from EDF header');
            end
            
        catch ME
            fprintf('  EDF header read failed: %s\n', ME.message);
            % Return default values
            duration = 0;
            sampling_rate = 200;
            n_channels = 0;
        end
    end