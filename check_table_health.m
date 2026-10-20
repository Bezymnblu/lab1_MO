function check_table_health(dataTable)
    % check_table_health - Задание 4: Проверка на наличие пустых ячеек (NaN)
    % с выводом точного физического адреса строки в файле Excel.

    fprintf('\n=== 🚨 НАЧАЛО ПРОВЕРКИ ЦЕЛОСТНОСТИ ДАННЫХ (Задание 4) ===\n');
    
    numCols = width(dataTable);
    colNames = dataTable.Properties.VariableNames;

    for j = 1:numCols
        currentColName = colNames{j};
        currentData = dataTable{:, j};
        
        % Находим логическую маску пустых ячеек (NaN)
        emptyMask = isnan(currentData);
        numMissing = sum(emptyMask);
        if numMissing > 0
            fprintf('⚠️  Столбец "%s": обнаружено %d пустых ячеек (NaN)!\n', currentColName, numMissing);
            
            emptyRows = find(emptyMask);
            
            % 🔥 ИДЕАЛЬНАЯ КАЛИБРОВКА: Индекс 1 равен строке №4 в Excel.
            % Формула: Номер строки в Excel = Индекс в MATLAB + 3.
            excelRows = emptyRows + 2;
            
            fprintf('   -> В Excel эти пустые ячейки находятся строго в строках: ');
            fprintf('%d ', excelRows);
            fprintf('\n');
        else
            fprintf('✅ Столбец "%s": в пределах рабочих данных пустых ячеек нет.\n', currentColName);
        end

    end
    
    fprintf('=== ✅ ПРОВЕРКА ДАННЫХ ЗАВЕРШЕНА ===\n');
end
