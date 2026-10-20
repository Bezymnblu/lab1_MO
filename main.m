% =========================================================================
% ЭТАЛОННОЕ ОБЪЕДИНЕНИЕ ЗАДАНИЙ 1, 2, 3, 4, 7 и 13 (main.m) - ЧИСТЫЕ ЧИСЛА
% =========================================================================
clear; clc; close all;

% Задание 12: Фиксируем seed для генератора случайных чисел
rng(42); 

% Имя твоего файла Excel в папке prj
datasetFile = 'Lab1_Dataset.xlsx'; 

fprintf('=== 📊 ЭТАП 1: ЗАДАНИЕ 1 (Чтение исходного файла) ===\n');
fullTable = read_tabular_data(datasetFile, false);
fprintf('✅ Исходный файл успешно прочитан.\n');

% Точный числовой срез Варианта 7 (строки 4-168 в Excel)
variant7Raw = fullTable(2:166, 13:14);
variant7Raw.Properties.VariableNames = {'x', 'y'};

% Переводим ячейки в числа
for i = 1:width(variant7Raw)
    if iscell(variant7Raw{:, i}) || isstring(variant7Raw{:, i})
        variant7Raw{:, i} = str2double(variant7Raw{:, i});
    end
end


fprintf('\n=== 🔍 ЭТАП 2: ЗАДАНИЕ 4 (Поиск пустых ячеек) ===\n');
check_table_health(variant7Raw);

fprintf('\n=== 🛠️  ЭТАП 3: ЗАДАНИЕ 7 (Интерактивная предобработка данных) ===\n');
variant7Processed = handle_missing_data(variant7Raw);


fprintf('\n=== 💾 ЭТАП 4: ЗАДАНИЕ 2 (Сохранение результатов из меню) ===\n');
outputFormat = 'csv';
savedFileName = write_tabular_data(variant7Processed, 'Variant7_Processed_Data', outputFormat);


%% 📈 ШАГ 5. Построение графика и скриншот (Задание 3 и 13)
fig = figure('Name', 'Вариант 7 - Обработанные данные', 'Color', 'w');

% Данные теперь гарантированно числовые, берем их напрямую!
plotX = variant7Processed.x;
plotY = variant7Processed.y;

% Находим рабочие строки без NaN для графика
cleanIdx = ~isnan(plotX) & ~isnan(plotY);
x_clean = plotX(cleanIdx);
y_clean = plotY(cleanIdx);
t_axis = (1:length(x_clean))';

% Строим две чистые числовые линии сигналов
plot(t_axis, x_clean, '-^', 'LineWidth', 1.5, 'Color', [0, 0.5, 0], 'MarkerFaceColor', [0, 0.5, 0]);
hold on;
plot(t_axis, y_clean, '-v', 'LineWidth', 1.5, 'Color', [0.7, 0, 0.5], 'MarkerFaceColor', [0.7, 0, 0.5]);

title('Раздел 3. График обработанных данных (Вариант 7)');
xlabel('Номер измерения (Индекс)');
ylabel('Значение сигналов');
legend({'Сигнал x', 'Сигнал y'}, 'Location', 'best');
grid on;
drawnow;

savedImageName = save_image_with_timestamp(fig, 'Variant7_Plot_Processed');


%% 🖨️  ШАГ 6. Полный вывод итоговой таблицы Варианта 7 (Задание 5)
fprintf('\n📊 ЗАДНИЕ 5 и 6: Полный вывод числового столбца Варианта 7.\n');
disp(variant7Processed);

%% 8️⃣ ШАГ 7. Конвертация таблицы в матрицу ячеек cells (Задание 8)
fprintf('\n📦 ЗАДАНИЕ 8: Конвертируем итоговую таблицу в матрицу ячеек (cells)...\n');

% 🔥 ГЛАВНАЯ КОМАНДА ЗАДАНИЯ 8:
% Функция table2cell берет объект table и превращает его в матрицу ячеек cell array.
variant7CellArray = table2cell(variant7Processed);

% Проверяем, какой тип данных получился (MATLAB должен написать 'cell')
fprintf('✅ Конвертация завершена успешно. Тип полученной переменной: "%s".\n', class(variant7CellArray));
fprintf('Размерность матрицы ячеек: %d строк на %d столбца.\n', size(variant7CellArray, 1), size(variant7CellArray, 2));

disp(variant7CellArray);
%%
fprintf('\n=== 📈 ЭТАП 4: ЗАДАНИЕ 9 (Выбор формулы и расчет x_new) ===\n');
% Чтобы формулы масштабирования считались корректно без поломок, 
% уберем NaN из таблицы перед расчетом
cleanTableForScale = rmmissing(variant7Processed);

% 🔥 ВЫЗЫВАЕМ НАШЕ НОВОЕ МЕНЮ МАСШТАБИРОВАНИЯ!
[variant7Scaled, scalingParams] = scale_data(cleanTableForScale);
disp(variant7Scaled);
%% 9️⃣ ШАГ 9. Восстановление изначального масштаба данных (Задание 10)
fprintf('\n=== 🔄 ЭТАП 5: ЗАДАНИЕ 10 (Восстановление изначального масштаба) ===\n');

% 🔥 ВЫЗЫВАЕМ НАШУ ФУНКЦИЮ ВОССТАНОВЛЕНИЯ!
% Мы передаем ей отмасштабированную таблицу variant7Scaled и структуру констант scalingParams.
% Программа сама поймет метод и посчитает обратную математику.
variant7Restored = inverse_scale_data(variant7Scaled, scalingParams);

% Выполняем Задание 5 для восстановленной таблицы: выводим её ПОЛНОСТЬЮ на экран!
fprintf('\n📊 Полный вывод восстановленной таблицы (Она должна один в один совпасть с числами до масштабирования):\n\n');
disp(variant7Restored);

% ПРОВЕРКА ТОЧНОСТИ (Инженерный контроль):
% Считаем максимальную разницу между восстановленными числами и числами до масштабирования.
% Она должна быть равна нулю (или ничтожно малой цифре вроде 1e-15 из-за округления процессора).
max_difference_x = max(abs(variant7Restored.x - cleanTableForScale.x));
max_difference_y = max(abs(variant7Restored.y - cleanTableForScale.y));
fprintf('📐 Контроль погрешности: максимальное расхождение по оси x = %g, по оси y = %g\n', max_difference_x, max_difference_y);
%% 🔟 ШАГ 10. Разбиение числового массива на 3 подмассива (Задание 11 и 12)
fprintf('\n=== ✂️  ЭТАП 6: ЗАДАНИЕ 11 и 12 (Разбиение выборки с фиксацией seed) ===\n');

% Переводим нашу итоговую обработанную числовую таблицу в чистую матрицу Num MATLAB
matrixForSplitting = table2array(variant7Scaled);

% Задаем пропорции строго по методичке в квадратных скобках (75% : 10% : 15%)
userRatios = [5 3 1];

% ВЫЗЫВАЕМ НАШУ ФУНКЦИЮ РАЗБИЕНИЯ
datasetSplits = split_dataset(matrixForSplitting, userRatios);

% =========================================================================
% 🔥 ПОЛНЫЙ ВЫВОД ВСЕХ РАЗДЕЛЕННЫХ ТАБЛИЦ С МЕТКАМИ (Задание 11)
% =========================================================================
fprintf('\n📺 ПОЛНЫЙ ВЫВОД РАЗДЕЛЕННЫХ МАССИВОВ С МЕТКАМИ:\n');

fprintf('\n--------------------------------------------------\n');
fprintf('  Метка: _train (Обучающая выборка)  \n');
fprintf('--------------------------------------------------\n');
% Выводим массив полностью
disp(datasetSplits.arr_train);

fprintf('\n--------------------------------------------------\n');
fprintf('  Метка: _val (Валидационная выборка)  \n');
fprintf('--------------------------------------------------\n');
% Выводим массив полностью
disp(datasetSplits.arr_val);

fprintf('\n--------------------------------------------------\n');
fprintf('  Метка: _test (Тестовая выборка)  \n');
fprintf('--------------------------------------------------\n');
% Выводим массив полностью
disp(datasetSplits.arr_test);

fprintf('=========================================================================\n');
fprintf('🎉 ПЕРВЫЕ ДВА РАЗДЕЛА ЛАБОРАТОРНОЙ РАБОТЫ (ЗАДАНИЯ 1-12) ПОЛНОСТЬЮ ГОТОВЫ!\n');
fprintf('=========================================================================\n');


