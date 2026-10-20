function splits = split_dataset(inputMatrix, ratios)
    % split_dataset - Задание 11 и 12: Разбиение числового массива на 3 части.
    % Входные параметры:
    %   inputMatrix - числовая матрица (Num MATLAB), полученная на Шаге 9
    %   ratios      - вектор из 3-х чисел (проценты, доли или соотношение)
    % Выходные параметры:
    %   splits      - структура с жесткими метками _train, _val, _test

    % === Шаг 1. Математическая нормировка вектора пропорций ===
    totalSum = sum(ratios);
    ratiosNorm = ratios / totalSum; % Теперь сумма элементов ВСЕГДА равна 1.0 (хоть 1:1:1, хоть 75/10/15)
    
    % Узнаем общее количество строк в числовом массиве
    numRows = size(inputMatrix, 1);
    
    % === Шаг 2. Расчет точного целочисленного количества строк для каждой выборки ===
    % Считаем размеры долями от общего количества
    numTrain = round(numRows * ratiosNorm(1));
    numVal = round(numRows * ratiosNorm(2));
    % Чтобы из-за округлений округления round не потерялась или не вылезла 1 строка, 
    % остаток жестко отдаем в тест:
    numTest = numRows - numTrain - numVal;
    
    % === Шаг 3. Контролируемое случайное перемешивание индексов (Задание 12) ===
    % Мы генерируем случайную перестановку индексов от 1 до numRows.
    % Благодаря фиксации rng(42) в main.m, этот порядок ВСЕГДА будет одинаковым!
    shuffledIndices = randperm(numRows);
    
    % === Шаг 4. Нарезаем индексы на 3 изолированные группы ===
    trainIdx = shuffledIndices(1 : numTrain);
    valIdx   = shuffledIndices(numTrain + 1 : numTrain + numVal);
    testIdx  = shuffledIndices(numTrain + numVal + 1 : end);
    
    % === Шаг 5. Формируем структуру со СТРОГИМИ метками методички (Задание 11) ===
    splits.arr_train = inputMatrix(trainIdx, :);
    splits.arr_val   = inputMatrix(valIdx, :);
    splits.arr_test  = inputMatrix(testIdx, :);
    
    % Выводим красивый пошаговый отчет в консоль для преподавателя
    fprintf('📐 [Задание 11] Массив из %d строк успешно разделен в соотношении [%.2f : %.2f : %.2f]:\n', ...
        numRows, ratiosNorm(1), ratiosNorm(2), ratiosNorm(3));
    fprintf('   -> Метка "_train" : %d строк\n', size(splits.arr_train, 1));
    fprintf('   -> Метка "_val"   : %d строк\n', size(splits.arr_val, 1));
    fprintf('   -> Метка "_test"  : %d строк\n', size(splits.arr_test, 1));
end
