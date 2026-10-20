function filename = write_tabular_data(dataTable, baseName, fileFormat)
    % write_tabular_data - Функция для сохранения таблиц с временной меткой.
    % Входные параметры:
    %   dataTable  - сама таблица с данными, которую хотим сохранить
    %   baseName   - желаемое имя файла (например, 'my_results')
    %   fileFormat - расширение файла (строка: 'csv', 'xlsx', 'txt' или 'mat')

    % Шаг 1: Получаем текущую дату и время на компьютере
    % Формат 'yyyyffffff_HHMMSS' сделает текст вида: '20260914_110500'
    currentTime = datetime('now', 'Format', 'yyyyMMdd_HHmmss');
    currentTimeStr = char(currentTime); % Превращаем время в обычный текст

    % Шаг 2: Собираем красивое итоговое имя файла
    % Получится, например: 'my_results_20260914_110500.csv'
    fileFormat = strrep(fileFormat, '.', ''); % Убираем точку, если пользователь её случайно ввёл
    filename = sprintf('%s_%s.%s', baseName, currentTimeStr, fileFormat);

    % Шаг 3: Сохраняем в зависимости от выбранного формата
    switch lower(fileFormat)
        case 'csv'
            % WriteVariableNames = true гарантирует, что шапка таблицы запишется в файл
            writetable(dataTable, filename, 'WriteVariableNames', true);
            
        case {'xlsx', 'xls'}
            writetable(dataTable, filename, 'WriteVariableNames', true);
            
        case 'txt'
            % Для текстовых файлов используем разделитель-табуляцию ('\t')
            writetable(dataTable, filename, 'Delimiter', '\t', 'WriteVariableNames', true);
            
        case 'mat'
            % Формат .mat не понимает таблицы напрямую в привычном виде,
            % поэтому мы конвертируем таблицу в структуру и сохраняем её командой save
            matData = table2struct(dataTable, 'ToScalar', true);
            save(filename, '-struct', 'matData');
            
        otherwise
            error('Ошибка: Направление сохранения в формат %s не поддерживается!', fileFormat);
    end
    
    fprintf('📁 Таблица успешно записана в файл: %s\n', filename);
end
