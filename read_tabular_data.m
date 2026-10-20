function dataTable = read_tabular_data(filePath, hasHeaders)
    % read_tabular_data - Универсальная функция чтения табличных данных.
    % Входные параметры:
    %   filePath   - путь к файлу (строка)
    %   hasHeaders - логическое значение (true, если есть заголовки, false - если нет)

    % Шаг 1: Извлекаем расширение файла (например, '.csv', '.xlsx')
    [~, ~, fileExt] = fileparts(filePath);
    fileExt = lower(fileExt); % переводим в нижний регистр для надежности

    % Шаг 2: Читаем файл в зависимости от его формата
    switch fileExt
        case '.csv'
            if hasHeaders
                dataTable = readtable(filePath);
            else
                % 'ReadVariableNames', false - заставляет MATLAB не искать шапку в файле
                dataTable = readtable(filePath, 'ReadVariableNames', false);
            end
            
        case {'.xlsx', '.xls'}
            if hasHeaders
                dataTable = readtable(filePath);
            else
                dataTable = readtable(filePath, 'ReadVariableNames', false);
            end
            
        case '.txt'
            % Для текстовых файлов указываем, что разделителем может быть пробел или табуляция
            if hasHeaders
                dataTable = readtable(filePath, 'FileType', 'text', 'Delimiter', 'space');
            else
                dataTable = readtable(filePath, 'FileType', 'text', 'Delimiter', 'space', 'ReadVariableNames', false);
            end
            
        case '.mat'
            % .mat - это родной формат MATLAB. Он хранит структуру переменных.
            matStruct = load(filePath);
            % Превращаем структуру в стандартную таблицу MATLAB таблицу
            dataTable = struct2table(matStruct);
            
        otherwise
            error('Ошибка: Формат файла %s не поддерживается!', fileExt);
    end

    % Шаг 3: Автоматическое именование, если заголовков в исходном файле не было (Задание 1)
    if ~hasHeaders
        numCols = width(dataTable); % Считаем, сколько всего столбцов в таблице
        newNames = cell(1, numCols); % Создаем пустой массив ячеек для новых имен
        
        for i = 1:numCols
            newNames{i} = sprintf('Column_%d', i); % Формируем имя 'Column_1', 'Column_2' ...
        end
        
        dataTable.Properties.VariableNames = newNames; % Присваиваем новые имена таблице
    end
end
