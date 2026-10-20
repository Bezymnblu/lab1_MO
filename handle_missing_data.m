function processedTable = handle_missing_data(inputTable)
    % handle_missing_data - Задание 7: Интерактивное меню предобработки данных
    % по пунктам методички.
    
    processedTable = inputTable;
    
    % Выводим красивое меню со строгими формулировками (Задание 7)
    fprintf('\n=======================================================\n');
    fprintf('   🛠️  МЕНЮ ПРЕДОБРАБОТКИ ДАННЫХ (Задание 7)          \n');
    fprintf('=======================================================\n');
    fprintf(' 1. Выбросить строки, содержащие некорректные данные\n');
    fprintf(' 2. Выполнить приведение типов\n');
    fprintf(' 3. Заполнить пустые ячейки средним значением по столбцу\n');
    fprintf(' 4. Заполнить пустые ячейки предыдущим или последующим значением\n');
    fprintf(' 5. Заполнить пустые ячейки нулями\n');
    fprintf(' 6. Заполнить ячейку указанным значением\n');
    fprintf('=======================================================\n');
    
    choice = input('👉 Введите номер действия (1-6): ');
    
    switch choice
        case 1
            % • Выбросить строки, содержащие некорректные данные
            processedTable = rmmissing(processedTable);
            fprintf('✅ Выполнено: Строки с пропусками и некорректными данными удалены.\n');
        case 2
            % • Выполнить приведение типов к общему вещественному виду (Исправлено)
            
            % Если на вход пришел текст или ячейки, принудительно переводим в числа.
            % Все целые числа станут double, а буквы/пропуски станут системными числовыми NaN double.
            if iscell(inputTable.x) || isstring(inputTable.x)
                processedTable.x = str2double(inputTable.x);
                processedTable.y = str2double(inputTable.y);
            else
                processedTable.x = double(inputTable.x);
                processedTable.y = double(inputTable.y);
            end
            
            % 🔥 ГЛАВНАЯ МАГИЯ MATLAB:
            % Настраиваем командное окно так, чтобы абсолютно все числа типа double 
            % на экране принудительно отображались в едином формате (4 знака после точки).
            % Теперь 1 превратится в 1.0000, но останется ЧИСЛОМ, без кавычек и скобок!
            format shortG; 
            
            fprintf('✅ Выполнено: Все целочисленные значения приведены к общему вещественному числовому типу double.\n');

        case 3
            % • Заполнить пустые ячейки средним значением по столбцу
            meanX = mean(processedTable.x, 'omitnan');
            meanY = mean(processedTable.y, 'omitnan');
            processedTable.x(isnan(processedTable.x)) = meanX;
            processedTable.y(isnan(processedTable.y)) = meanY;
            fprintf('✅ Выполнено: Пустые ячейки заполнены средними по столбцам (x=%.4f, y=%.4f).\n', meanX, meanY);
            
        case 4
            % • Заполнить пустые ячейки предыдущим или последующим значением
            fprintf('   Выберите метод заполнения:\n');
            fprintf('   1 - Предыдущим значением (ffill)\n');
            fprintf('   2 - Последующим значением (bfill)\n');
            subChoice = input('   👉 Ваш выбор (1 или 2): ');
            if subChoice == 1
                processedTable = fillmissing(processedTable, 'previous');
                fprintf('✅ Выполнено: Заполнено предыдущими значениями.\n');
            else
                processedTable = fillmissing(processedTable, 'next');
                fprintf('✅ Выполнено: Заполнено последующими значениями.\n');
            end
            
        case 5
            % • Заполнить пустые ячейки нулями
            processedTable = fillmissing(processedTable, 'constant', 0);
            fprintf('✅ Выполнено: Все пустые ячейки заполнены нулями.\n');
            
        case 6
            % • Заполнить ячейку указанным значением
            customValue = input('⌨️  Введите ваше указанное значение: ');
            processedTable = fillmissing(processedTable, 'constant', customValue);
            fprintf('✅ Выполнено: Все пустые ячейки заполнены значением %.4f.\n', customValue);
            
        otherwise
            warning('⚠️  Неверный выбор! Данные не изменены.');
    end
end
