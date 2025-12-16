# coding: cp1251

if 'const':
  et3dNo3dDocument  =-7;  # "Документ не активизирован или не является 3D-моделью"
  et3dAbort         =-1;  # "Аварийное завершение (?)"
  et3dError1        = 1;  # "3D документ уже создан"
  et3dError2        = 2;  # "Документ 3D-модели уже открыт"
  et3dError3        = 3;  # "Ошибка при создании документа 3D-модели"
  et3dError4        = 4;  # "Документ 3D-модели должен быть сборкой"
  et3dError5        = 5;  # "Объект не найден в данном документе 3D-модели"
  et3dError6        = 6;  # "Неверный тип параметров редактирования"
  et3dError7        = 7;  # "Объект должен быть 3D локальной системой координат"
  et3dError8        = 8;  # "Недостаточное количество эскизов для операции"
  et3dError9        = 9;  # "У контура слишком много осей"
  et3dError10       = 10; # "У контура не хватает осей"
  et3dError11       = 11; # "Слишком много контуров"
  et3dError12       = 12; # "У контура есть "звезда"
  et3dError13       = 13; # "Самопересечение контура"
  et3dError14       = 14; # "Самопересечение в продолжении контура"
  et3dError15       = 15; # "Пересечение контуров"
  et3dError16       = 16; # "Образующий контур не пересекает тела"
  et3dError17       = 17; # "Образующий контур не пересекает сечения (для операции построения тела по сечениям)"
  et3dError18       = 18; # "Фрагмент не найден в библиотеке"
  et3dError19       = 19; # "Контур должен быть замкнут"
  et3dError20       = 20; # "Контур должен быть разомкнут"
  et3dError21       = 21; # "Пересечение контура с осью"
  et3dError22       = 22; # "Вырожденная ось у контура"
  et3dError23       = 23; # "Тело состоит из отдельных частей"
  et3dError24       = 24; # "Ошибочная кривая"
  et3dError25       = 25; # "Ошибочный контур"
  et3dError26       = 26; # "Ошибочная поверхность"
  et3dError27       = 27; # "Ошибочное тело"
  et3dError28       = 28; # "Ошибочный параметр"
  et3dError29       = 29; # "Неправильно задана толщина"
  et3dError30       = 30; # "Не последовательное расположение сечений вдоль кривой (для операции построения тела по сечениям)"
  et3dError31       = 31; # "Объект самопересекается"
  et3dError32       = 32; # "Объекты не пересекаются"
  et3dError33       = 33; # "Объекты пересекается с ошибкой"
  et3dError34       = 34; # "Ошибка в булевой операции"
  et3dError35       = 35; # "Ребра не найдены"
  et3dError36       = 36; # "Ошибка при подготовке операции"
  et3dError37       = 37; # "Ошибка при создании фаски ребра"
  et3dError38       = 38; # "Ошибка при создании скругления ребра"
  et3dError39       = 39; # "Созданы фаски не на всех ребрах"
  et3dError40       = 40; # "Скруглены не все ребра"
  et3dError41       = 41; # "Ошибка при создании поверхности фаски ребра"
  et3dError42       = 42; # "Ошибка при создании поверхности скругления ребра"
  et3dError43       = 43; # "Слишком большые катеты фаски"
  et3dError44       = 44; # "Слишком большой радиус скругления"
  et3dError45       = 45; # "Фаски построены не для всех ребер"
  et3dError46       = 46; # "Скруглены не все ребра"
  et3dError47       = 47; # "Ошибка резки поверхностью"
  et3dError48       = 48; # "Ошибка при создании тонкостенного тела"
  et3dError49       = 49; # "Слишком большая толщина стенки при создании тонкостенного тела"
  et3dError50       = 50; # "Ошибочная грань"
  et3dError51       = 51; # "Контур пустой"
  et3dError52       = 52; # "Неизвестная ошибка постановки ребра жесткости"
  et3dError53       = 53; # "Неизвестная ошибка уклона граней тела"
  et3dError54       = 54; # "Неизвестная ошибка"
  et3dError55       = 55; # "Ошибка в определении имени файла"
  et3dError56       = 56; # "Ошибка в определении детали"
  et3dError57       = 57; # "Ошибка в определении объекта"
  et3dError58       = 58; # "Ошибка в определении типа сопряжения"
  et3dError59       = 59; # "Ошибка при выполненин сопряжения"
  et3dError60       = 60; # 
  et3dError61       = 61; # "Нужно выполнить процесс по выделению объектов"
  et3dError62       = 62; # "Ошибка в определении массива объектов"
  et3dError63       = 63; # "Массив объектов не соответствует документу 3D-модели детали"
  et3dError64       = 64; # "Тип 3D объекта задан неверно"
  et3dError65       = 65; # "Ошибка при вставке 3D-модели детали"
  et3dError66       = 66; #   66 Нужно завершить режим разнесенного вида
  et3dError67       = 67; #   67 Базовая операция может быть только одна
  et3dError68       = 68; #   68 Для построения нужна базовая операция
  et3dError69       = 69; #   69 Попытка изменить редактируемую деталь в процессе вставки детали в модель
  et3dError70       = 70; #   70 Ошибка! Попытка поставить в макро \n\r недопустимый объект"
  et3dError71       = 71; #   71 Ошибка! Библиотека фрагментов уже закрыта или не открывалась"
  et3dError72       = 72; #   72 Ошибка! Библиотека фрагментов уже открыта"
  et3dError73       = 73; #   73 Ошибка! Файл библиотеки фрагментов не найден"
  et3dError74       = 74; #   74 Ошибка в структуре файла библиотеки фрагментов"
  et3dError75       = 75; #   75 Ошибка в имени файла библиотеки фрагментов"
  et3dError76       = 76; #   76 Ошибка в имени фрагмента для библиотеки фрагментов"
  et3dError77       = 77; #   77 Ошибка! Доступ к фрагменту в библиотеке фрагментов  невозможен"
  et3dError78       = 78; #   78 Ошибка! в выбранном направлении отсутствует поверхность
  et3dError79       = 79; #   79 Ошибка! локальное тело поглощает результат
  et3dError80       = 80; #   80 Ошибка! ломаная должна иметь не менее двух вершин
  et3dError81       = 81; #   81 Ошибка! сплайн должен иметь не менее двух вершин
  et3dError82       = 82; #   82 Ошибка в определении эскиза
  et3dError83       = 83; #   83 Ошибка в определении ребра
  et3dError84       = 84; #   84 Ошибка в определении вершины
  et3dError85       = 85; #   85 Ошибка в определении планарной плоскости
  et3dError86       = 86; #   86 Ошибка в определении поверхности
  et3dError87       = 87; #   87 Ошибка создания файла библиотекаря моделей
  et3dError88       = 88; #   88 "Неподходящий тип объекта для проецирования"
  et3dError89       = 89; #   89 вырожденная проекция объекта на плоскость эскиза
  et3dError90       = 90; #   90 некоторые ребра не спроецированы
  et3dError91       = 91; #   91 "Документ с таким именем уже открыт"
  et3dError92       = 92; #   92 "Компонент не найден"
  et3dError93       = 93; #   93 "Материал у компонента не найден"
  et3dError94       = 94; #   94 Попытка изменить деталь из библиотеки моделей или стандартную деталь
  et3dError95       = 95; #   95 Компонент не является компонентом сборки
  et3dError96       = 96; #   96 Недопустимая операция для компонента, редактируемого на месте
  et3dError97       = 97; #   97 Цвет и оптические свойства объекта в сборке изменить нельзя
  et3dError98       = 98;  #  98 Импортированная поверхность задана неверно
  et3dError99       = 99;  #  99 Геометрия модели изменилась, данные могут быть не корректними
  et3dError100      = 100;  # 100 Ошибка при получении данных триангуляции
  et3dError101      = 101;  # 101 Ошибка при получении данных цилов
  et3dError102      = 102;  # 102 Невыбран ни один плоский объект
  et3dError103      = 103;  # 103 Ошибка при получении данных поверхности
  et3dError104      = 104;  # 104 Ошибка при получении данных 3d плоскости
  et3dError105      = 105;  # 105 Ошибка при получении данных конической поверхности
  et3dError106      = 106;  # 106 Ошибка при получении данных цилиндлической поверхности
  et3dError107      = 107;  # 107 Ошибка при получении данных сферы
  et3dError108      = 108;  # 108 Ошибка при получении данных тора
  et3dError109      = 109;  # 109 Ошибка при получении данных кривой
  et3dError110      = 110;  # 110 Ошибка при получении данных линии
  et3dError111      = 111;  # 111 Ошибка при получении данных окружности
  et3dError112      = 112;  # 112 Ошибка при получении данных эллипса
  et3dError113      = 113;  # 113 "У 3d модели нет размеров листа"

  MAXERROR3D        = 113;

# типы направлений Direction_Type
  dtNormal = 0;      # прямое направление
  dtReverse = 1;     # обратное направление
  dtBoth = 2;        # в обе стороны
  dtMiddlePlane = 3; # средняя плоскость

# типы операций End_Type
  etBlind = 0;           # строго на глубину
  etThroughAll = 1;      # насквозь всей детали
  etUpToVertexTo = 2;    # до вершины  до
  etUpToVertexFrom = 3;  # до вершины  за
  etUpToSurfaceTo = 4;   # до грани  до
  etUpToSurfaceFrom = 5; # до грани  за
  etUpToNearSurface = 6; # до ближайшей грани

# типы Part_Type
  pInPlace_Part = -4 # деталь без ссылки на источник
  pNew_Part  = -3;  # новая деталь
  pEdit_Part = -2;  # редактируемая деталь
  pTop_Part  = -1;  # верхний компонент

# типы сопряжений MateConstraintType
  mc_Coincidence   = 0; # совпадение/принадлежность объектов
  mc_Parallel      = 1; # параллельность
  mc_Perpendicular = 2; # перпендикулярность
  mc_Tangency      = 3; # касательность
  mc_Concentric    = 4; # концентричность
  mc_Distance      = 5; # расстояние между объектами
  mc_Angle         = 6; # угол между объектами
  mc_InPlace       = 7; # точное совпадение двух плоскостей

# режим визуализации модели ViewMode
  vm_Wireframe     = 0; # каркас
  vm_HiddenRemoved = 1; # удаление невидимых линий
  vm_HiddenThin    = 2; # невидимые линии тонкие
  vm_Shaded        = 3; # полутоновой

# типы предопределенных проекций отображения  ProjectionType
  vp_None          = -1; # Не определена (ни то, ни се)
  vp_NormalTo      = 0;  # Нормально к текущему планару
  vp_Front         = 1;  # Спереди  - Фронтальная плоскость
  vp_Rear          = 2;  # Сзади
  vp_Up            = 3;  # Сверху   - Горизонтальная плоскость
  vp_Down          = 4;  # Снизу
  vp_Left          = 5;  # Слева    - Профильная плоскость
  vp_Right         = 6;  # Справа
  vp_IsoXYZ        = 7;  # Изометрия XYZ
  vp_IsoYZX        = 8;  # Изометрия YZX
  vp_IsoZXY        = 9;  # Изометрия ZXY
  vp_Dio           = 10; # Диметрия


# типы Obj3dType
#------------------------------------------------------------------------------
# любые изменения в этом enum'е необходимо учитывать в D3LtEntityCollection::IsOurEntity
# ---
  o3d_unknown                    = 0;   # неизвестный (включает все объекты)

  # default'ные элементы
  o3d_planeXOY                   = 1;   # плоскость XOY
  o3d_planeXOZ                   = 2;   # плоскость XOZ
  o3d_planeYOZ                   = 3;   # плоскость YOZ
  o3d_pointCS                    = 4;   # точка начала системы координат

  # элементы детали
  o3d_sketch                     = 5;   # эскиз
  o3d_face                       = 6;   # поверхность
  o3d_edge                       = 7;   # грань
  o3d_vertex                     = 8;   # вершина

  # конструктивные элементы
  o3d_axis2Planes                = 9;   # ось по двум плоскостям
  o3d_axis2Points                = 10;  # ось по двум точкам
  o3d_axisConeFace               = 11;  # ось конической грани
  o3d_axisEdge                   = 12;  # ось проходящая через ребро
  o3d_axisOperation              = 13;  # ось операции
  o3d_planeOffset                = 14;  # смещённая плоскость
  o3d_planeAngle                 = 15;  # плоскость под углом
  o3d_plane3Points               = 16;  # плоскость по 3-м точкам
  o3d_planeNormal                = 17;  # нормальная плоскость
  o3d_planeTangent               = 18;  # касательная плоскость
  o3d_planeEdgePoint             = 19;  # плоскость через ребро и вершину
  o3d_planeParallel              = 20;  # плоскость через вершину параллельно другой плоскости
  o3d_planePerpendicular         = 21;  # плоскость через вершину перпендикулярно ребру
  o3d_planeLineToEdge            = 22;  # плоскость через ребро пар-но/пер-но другому ребру
  o3d_planeLineToPlane           = 23;  # плоскость через ребро пар-но/пер-но грани

  # операции
  o3d_baseExtrusion              = 24;  # базовая операция выдавливания
  o3d_bossExtrusion              = 25;  # приклеивание выдавливанием
  o3d_cutExtrusion               = 26;  # вырезать выдавливанием
  o3d_baseRotated                = 27;  # базовая операция вращения
  o3d_bossRotated                = 28;  # приклеивание вращением
  o3d_cutRotated                 = 29;  # вырезать вращением
  o3d_baseLoft                   = 30;  # базовая операция вращения
  o3d_bossLoft                   = 31;  # приклеивание вращением
  o3d_cutLoft                    = 32;  # вырезать вращением
  o3d_chamfer                    = 33;  # операция "фаска"
  o3d_fillet                     = 34;  # операция "скругления"
  o3d_meshCopy                   = 35;  # операция копирования по сетке
  o3d_circularCopy               = 36;  # операция копирования по концентрической сетке
  o3d_curveCopy                  = 37;  # операция копирования по кривой
  o3d_circPartArray              = 38;  # операция массив по концентрической сетке для сборки
  o3d_meshPartArray              = 39;  # операция массив по сетке для сборки
  o3d_curvePartArray             = 40;  # операция массив по кривой для сборки
  o3d_derivPartArray             = 41;  # операция массив по кривой для сборки
  o3d_incline                    = 42;  # операция "уклон"
  o3d_shellOperation             = 43;  # операция "оболочка"
  o3d_ribOperation               = 44;  # операция "ребро жесткости"
  o3d_baseEvolution              = 45;  # кинематическая операция
  o3d_bossEvolution              = 46;  # приклеинть кинематически
  o3d_cutEvolution               = 47;  # вырезать кинематически
  o3d_mirrorOperation            = 48;  # операция "зеркальная копия"
  o3d_mirrorAllOperation         = 49;  # операция "зеркально отразить все"
  o3d_cutByPlane                 = 50;  # операция "сечение поверхностью"
  o3d_cutBySketch                = 51;  # операция "сечение эскизом"
  o3d_holeOperation              = 52;  # отверстие

  # кривые
  o3d_polyline                   = 53;  # ломаная
  o3d_conicSpiral                = 54;  # Коническая спираль
  o3d_spline                     = 55;  # сплайн
  o3d_cylindricSpiral            = 56;  #
  o3d_importedSurface            = 57;  # импортирванная поверхность
  o3d_thread                     = 58;  # ПК [4/10/2003] Условное отображение резьбы

  o3d_EvolutionSurface           = 59;  # Кинематическая поверхность
  o3d_ExtrusionSurface           = 60;  # Поверхность выдавливания
  o3d_RotatedSurface             = 61;  # Поверхность вращения
  o3d_LoftSurface                = 62;  # Поверхность по сечениям
  o3d_MacroObject                = 63;  # Макрообъект 3D
  o3d_UnionComponents            = 64;
  o3d_MoldCavity                 = 65;
  o3d_planeMiddle                = 66;
  o3d_controlPoint               = 67;  # Контрольная точка
  o3d_conjunctivePoint           = 68;  # Присоединительная точка
  o3d_aggregate                  = 69;  

# default'ные элементы
  o3d_point3D                    = 70;  # точка 3D
  o3d_axisOX                     = 71;  # ось OX
  o3d_axisOY                     = 72;  # ось OY
  o3d_axisOZ                     = 73;  # ось OZ
  
  o3d_sheetMetalBody             = 74;
  o3d_sheetMetalBend             = 75;
  o3d_sheetMetalLineBend         = 76;
  o3d_sheetMetalHole             = 77;
  o3d_sheetMetalCut              = 78;
  o3d_UnHistoried                = 79;
  o3d_baselineDimension3D        = 80;
  o3d_lineDimension3D            = 81;
  o3d_radialDimension3D          = 82;
  o3d_diametralDimension3D       = 83;
  o3d_angleDimension3D           = 84;
  o3d_localCoordinateSystem      = 85;
  o3d_leader3D                   = 86;
  o3d_markLeader3D               = 87;
  o3d_rough3D                    = 88;
  o3d_positionLeader3D           = 89;
  o3d_brandLeader3D              = 90;
  o3d_base3D                     = 91;
  o3d_tolerance3D                = 92;
  o3d_SplitLine                  = 93;
  o3d_SurfacePatch               = 94;
  o3d_FaceRemover                = 95;
  o3d_SurfaceSewer               = 96;
  o3d_NurbsSurface               = 97;
  o3d_SurfacesIntersectionCurve  = 98;

  o3d_lastEntityElement          = 99;  # Всегда последний из Entity!!!

  # элементы не являющиеся Entity
  o3d_variable                   = 100; # параметрическая переменная
  o3d_placement                  = 101; # местоположение
  o3d_entityCollection           = 102; # массив 3d объектов
  o3d_document                   = 103; # документ 3d
  o3d_part                       = 104; # деталь
  o3d_entity                     = 105; # объект
  o3d_mateConstraint             = 106; # сопряжение
  o3d_mateConstraintCollection   = 107; # массив сопряжений
  o3d_partCollection             = 108; # массив элементов сборки

  # объединённые типы объектов для создания EntityCollection
  o3d_constrElement              = 109; # конструктивные элементы - плоскости и оси (конструктивные от o3d_axis2Planes до o3d_plane3Points)
  o3d_operationElement           = 110; # операции (от o3d_baseExtrusion до o3d_fillet)
  o3d_curveElement               = 111; # кривые (пространственные и ребра)

  o3d_rasterFormat               = 112; # интерфейс параметров для конвертации в растровый формат
  o3d_additionFormat             = 113; # интерфейс параметров для конвертации в дополнительные форматы jgs, sat,xt,x_b, step, stl, VRML

  o3d_bodyCollection             = 114; # интерфейс массива тел 3D
  o3d_body                       = 115; # интерфейс тела 3D
  o3d_faceCollection             = 116; # интерфейс массива граней
  o3d_tessellation               = 117; # интерфейс триангуляции
  o3d_facet                      = 118; # интерфейс триангуляционной пластины
  o3d_featureCollection          = 119; # интерфейс массива объектов дерева
  o3d_feature                    = 120; # интерфейс объекта дерева
  o3d_edgeCollection             = 121; # интерфейс массива ребер
  o3d_orientedEdge               = 122; # интерфейс ориентированного ребра
  o3d_orientedEdgeCollection     = 123; # интерфейс массива ориентированных ребер
  o3d_loop                       = 124; # интерфейс цикла
  o3d_loopCollection             = 125; # интерфейс массива циклов
  o3d_curve3D                    = 126; # интерфейс математической кривой в трехмерном пространстве
  o3d_surface                    = 127; # интерфейс математической кривой в трехмерном пространстве
  o3d_massInertiaParam           = 128; # Интерфейс параметров для расчета массо-центровочных характеристик
  o3d_lineseg3dParam             = 129; # Интерфейс параметров 3d LineSeg
  o3d_circle3dParam              = 130; # Интерфейс параметров 3d Circle
  o3d_ellipse3dParam             = 131; # Интерфейс параметров 3d Ellipce
  o3d_nurbsPoint3dParam          = 132; # Интерфейс параметров точки для Nurbs 3D
  o3d_nurbsPoint3dCollection     = 133; # Интерфейс массива точек для Nurbs 3D
  o3d_nurbsPoint3dCollCollection = 134; # Интерфейс массива массивов точек для Nurbs 3D Surface
  o3d_nurbsKnotCollection        = 135; # Интерфейс массива узлов для Nurbs 3D
  o3d_nurbs3dParam               = 136; # Интерфейс параметров Nurbs-сплайна 3D
  o3d_planeParam                 = 137; # Интерфейс параметров плоскости
  o3d_coneParam                  = 138; # Интерфейс параметров конической поверхности
  o3d_cylinderParam              = 139; # Интерфейс параметров цилиндрической поверхности
  o3d_sphereParam                = 140; # Интерфейс параметров сферы
  o3d_torusParam                 = 141; # Интерфейс параметров тора
  o3d_nurbsSurfaceParam          = 142; # Интерфейс параметров Nurbs-поверхности
  o3d_mateConstraintGroup        = 143; # Объект дерева :группа сопряжений
  o3d_measurer                   = 144; # Интерфейс для измерений расстояния и угла между двумя примитивами ( гранями, ребрами, вершинами)
  o3d_selectionMng               = 145; # Интерфейс менеджера селектированных объектов
  o3d_chooseMng                  = 146; # Интерфейс менеджера выбора (подсветки) объектов
  o3d_arc3dParam                 = 147; # Интерфейс параметров 3d Arc
  o3d_deletedCopyCollection      = 148; # Интерфейс массива удаленных индексов для оперций копирования и массивов компонент
  
  o3d_viewProjection             = 149;
  o3d_viewProjectionCollection   = 150;
  o3d_attribute                  = 151;
  o3d_attributeCollection        = 152;
  o3d_componentPositioner        = 153;
  o3d_modelLibrary               = 154;
  o3d_ObjectsFilter3D            = 155;
  o3d_coordinate3dCollection     = 156;
  o3d_intersectionResult         = 157;
  o3d_PolygonalLineVertexParam   = 158;
  o3d_variableCollection         = 159;
  o3d_sTrackingPointsMeasurer    = 160;
  o3d_surfaceElement             = 161;
  o3d_designationElement         = 162;
  o3d_copyleftObject             = 163;
  o3d_firstEntityElement2        = 500;
  o3d_Equidistant3D              = 501;
  o3d_TrimmedCurve               = 502;
  o3d_TrimmedCurveObject         = 503;
  o3d_AuxMeshCopy                = 504;
  o3d_AuxCircularCopy            = 505;
  o3d_AuxCurveCopy               = 506;
  o3d_PointDrivenPattern         = 507;
  o3d_PartsPointDrivenPattern    = 508;
  o3d_AuxMirrorOperation         = 509;
  o3d_ConnectCurve               = 510;
  o3d_ConnectCurveObject         = 511;
  o3d_FilletCurve                = 512;
  o3d_FilletCurveObject          = 513;
  o3d_EquidistantSurface         = 514;
  o3d_RuledSurface               = 515;
  o3d_TrimmedSurface             = 516;
  o3d_ExtensionSurface           = 517;
  o3d_SurfaceThickening          = 518;
  o3d_Arc3D                      = 519;
  o3d_AuxPointDrivenPattern      = 520;
  o3d_BodiesPointDrivenPattern   = 521;
  o3d_TablePattern               = 522;
  o3d_PartsTablePattern          = 523;
  o3d_AuxTablePattern            = 524;
  o3d_BodiesTablePattern         = 525;
  o3d_MeshPointsSurface          = 526;
  o3d_CloudPointsSurface         = 527;
  o3d_BodiesMeshCopy             = 528;
  o3d_BodiesCircularCopy         = 529;
  o3d_BodiesCurveCopy            = 530;
  o3d_Scaling3D                  = 531;
  o3d_lastEntityElement2         = 1500

#------------------------------------------------------------------------------
#определения для конвертации в дополнительные форматы jgs, sat, xt, step, stl, VRML
# ---
# enum D3FormatConvType
  format_SAT   = 1;
  format_XT    = 2;
  format_STEP  = 3;
  format_IGES  = 4;
  format_VRML  = 5;
  format_STL   = 6;

#------------------------------------------------------------------------------
# Типы используемого цвета
# enum UseColor
  useColorUnknown = -1; # тип не определен
  useColorOur     =  0; # собственный цвет
  useColorOwner   =  1; # цвет хозяина
  useColorSource  =  2; # цвет источника

#------------------------------------------------------------------------------
# Тип перемещения
# enum Positioner_Type
  pnMove        = 0;
  pnRotate      = 1;

#------------------------------------------------------------------------------
# Типы пересечений.
# enum Intersection_Type
  itTangentPoint        = 1;
  itTangentCurve        = 2;
  itTangentSurface      = 3;
  itBody                = 4;

#------------------------------------------------------------------------------
# Типы математических объектов учавствующих в сопряжении.
# enum MateType
  ksMateUnknown = 0;
  ksMatePoint   = 1;
  ksMateLine    = 2;
  ksMatePlane   = 3;
  ksMateCylinder= 4;
  ksMateCone    = 5;
  ksMateSphere  = 6;
  ksMateTorus   = 7;
  ksMateCircle  = 8;

#------------------------------------------------------------------------------
# Типы действий для оперций над телами
# enum ChooseBodiesType
  ksNewBody               = 0;
  ksAutomaticDefinition   = 1;
  ksManualEditing         = 2;
  ksAllBodies             = 3;

#------------------------------------------------------------------------------
# Типы булевых операций над твердыми телами
# enum BooleanType
  ksBooleanUnknown      = 0;
  ksIntersect           = 1;
  ksDifference          = 2;
  ksUnion               = 3;

#------------------------------------------------------------------------------
# Способ построения сегмента ломоной.
# enum LineBuildingType
  ksLBTByPoint          = 0;
  ksLBTXDirection       = 1;
  ksLBTYDirection       = 2;
  ksLBTZDirection       = 3;
  ksLBTParallel         = 4;
  ksLBTPerpendicular    = 5;
  ksLBTByPoint3DParams  = 6;

#------------------------------------------------------------------------------
# Способ построения сегмента ломоной.
# enum Part7CollectionTypeEnum
  ksAllParts    = 0;
  ksUniqueParts = 1;

#------------------------------------------------------------------------------
# Способ определения длины развертки.
# enum ksUnfoldTypeEnum
  ksCoefficient         = 0;
  ksValueBend           = 1;
  ksDecreaseBend        = 2;
  ksTableBends          = 3;

#------------------------------------------------------------------------------
# Тип размещения сгиба на ребре .
# enum ksBendDisposalEnum
  ksBendDisposalAllLength       = 0;
  ksBendDisposalCentre          = 1;
  ksBendDisposalLeft            = 2;
  ksBendDisposalRight           = 3;
  ksBendDisposalTwo             = 4;
  ksBendDisposalLeftAndWidth    = 5;
  ksBendDisposalRightAndWidth   = 6;

#------------------------------------------------------------------------------
# Тип определения длины.
# enum ksBendLengthTypeEnum
  ksBendLengthByContinue        = 0;
  ksBendLengthByContour         = 1;
  ksBendLengthByTouch           = 2;

#------------------------------------------------------------------------------
# Тип смещения.
# enum ksBendOffsetTypeEnum
  ksBendOffsetIn                = 0;
  ksBendOffsetOut               = 1;
  ksBendOffsetLineOutside       = 2;
  ksBendOffsetLineInside        = 3;
  ksBendOffsetByTouch           = 4;

#------------------------------------------------------------------------------
# Тип построения боковой стороны сгиба.
# enum ksBendSideTypeEnum
  ksBendSideByAngle     = 0;
  ksBendSideByWidening  = 1;

#------------------------------------------------------------------------------
# Тип освобождения cгиба.
# enum ksBendReleaseTypeEnum
  ksBendReleaseByRect   = 0;
  ksBendReleaseByCircle = 1;

#------------------------------------------------------------------------------
# Способ освобождения угла сгиба.
# enum ksBendAngleReleaseTypeEnum
  ksBendAngleBendOnly   = 0;
  ksBendAngleIn         = 1;
  ksBendAngleAllBends   = 2;

#------------------------------------------------------------------------------
# Способ сгиба.
# enum ksBendTypeEnum
  ksLineBend            = 0;
  ksBendLineOutside     = 1;
  ksBendLineInside      = 2;
  ksBendByTouch         = 3;

#------------------------------------------------------------------------------
# Тип построения отверстия и выреза .
# enum ksHoleCutTypeEnum
  ksHoleCutByWidth      = 0;
  ksHoleCutByDepth      = 1;
  ksHoleCutUpToSurface  = 2;

#-----------------------------------------------------------------------------
# Способ определения области применения для компонентов в сборочной операции
# enum ChoosePartsType
  ksChAutomaticDefinition       = 1;
  ksChManualEditing             = 2;
  ksChAllParts                  = 3;
  ksChNoLibraryParts            = 4;

#-----------------------------------------------------------------------------
# Область применения
# enum ChooseType
  ksChBodiesAndParts    = 1;
  ksChParts             = 2;
  ksChBodies            = 3;

#------------------------------------------------------------------------------
# Способы построения пространственной точки.
# enum ksPoint3DTypeEnum
  ksPUnknown            = 0;
  ksPParamCoord         = 1;
  ksPDisplace           = 2;
  ksPIntersect          = 3;
  ksPCenter             = 4;
  ksPCurve              = 5;
  ksPSurface            = 6;
  ksPProjection         = 7;
  ksPCylindrCoord       = 8;
  ksPSphericCoord       = 9;

#------------------------------------------------------------------------------
# Типы смещений при способе построения точки вдоль кривой.
# enum ksPoint3DCurveParamTypeEnum
  ksOffsetByU           = 1;
  ksOffsetByLen         = 2;
  ksOffsetByAngle       = 3;

#------------------------------------------------------------------------------
# Типы смещений при способе построения точки на поверхности.
# enum ksPoint3DSurfaceParamTypeEnum
  ksOffsetByUV          = 1;
  ksOffsetByLenFromObj  = 2;

#------------------------------------------------------------------------------
# Результат измерения расстояния и угла между поверхностями.
# enum ksMeasureResultEnum
  ksMResUnknown                 = 0;
  ksMResAxisAxisCoaxial         = 1;
  ksMResAxisAxisParallel        = 2;
  ksMResAxisAxisIntersect       = 3;
  ksMResAxisAxisDistant         = 4;
  ksMResAxisSurfColinear        = 5;
  ksMResAxisSurfParallel        = 6;
  ksMResAxisSurfIntersect       = 7;
  ksMResAxisSurfDistant         = 8;
  ksMResSurfSurfColinear        = 9;
  ksMResSurfSurfParallel        = 10;

#------------------------------------------------------------------------------
# Тип ориентирования ЛСК.
# enum ksOrientationTypeEnum
  ksAxisOrientation     = 0;
  ksEulerCorners        = 1;
  ksOrientByObject      = 2;

#------------------------------------------------------------------------------
# Тип параметров объекта.
# enum ksModelObjectParamTypeEnum
  ksMOAllParam                  = 1;
  ksMOPartAllParam              = 2;
  ksMOCurrentLSKAllParam        = 3;

#------------------------------------------------------------------------------
# Стили 3D линий для отрисовки с помощью OpenGL.
# enum ks3DLineStyle
  ksCS3DNoDrawing               = 0;
  ksCS3DSolid                   = 1;
  ksCS3DDashed                  = 2;
  ksCS3DDotted                  = 3;
  ksCS3DDashDot                 = 4;
  ksCS3DDashDotLDash2Dots       = 5;

#------------------------------------------------------------------------------
# Тип загрузки компонента.
# enum ksLoadStateEnum
  ksLUnknown    = -1;
  ksLCompletely = 0;
  ksLUnload     = 1;
  ksLPartially  = 2;

#------------------------------------------------------------------------------
# Режим фильтрации отображаемых граней внешнего объекта.
# enum ksFacetCullingMode
  ksFSMNone     = 0;
  ksFSMFront    = 1;
  ksFSMBack     = 2;
  ksFSMAll      = 3;

#------------------------------------------------------------------------------
# Обход углов эквидистанты 3D.
# enum ksEquidistant3DCutModeEnum
  ksECMUnknown  = 0;
  ksECMLineSeg  = 1;
  ksECMCircle   = 2;

#------------------------------------------------------------------------------
# Способ задания базовой точки.
# enum ksPatternBasePointTypeEnum
  ksCRAuto              = 0;
  ksCRManual            = 1;
  ksCRFirstObject       = 2;

#------------------------------------------------------------------------------
# Тип продления поверхности.
# enum ksExtensionSurfaceTypeEnum
  ksESTUnknown          = -1;
  ksESTSelf             = 0;
  ksESTTangent          = 1;
  ksESTDirection        = 2;

#------------------------------------------------------------------------------
# Способ ограничения.
# enum ksExtensionLimitTypeEnum
  ksETLUnknown  = -1;
  ksETLength    = 0;
  ksETLVertex   = 1;

#------------------------------------------------------------------------------
# Типы параметров вектора.
# enum ksVector3DParametersTypeEnum
  ksVector3DUnknown             = 0;
  ksVector3D2Vertex             = 1;
  ksVector3DCSAngle             = 2;
  ksVector3DAxis                = 3;
  ksVector3DCoefficients        = 4;
  ksVector3D2Angles             = 5;
  ksVector3DEdge                = 6;
  ksVector3DPlane               = 7;
  ksVector3DSurface             = 8;
  ksVector3DCurve               = 9;
  ksVector3DScreen              = 10;

#------------------------------------------------------------------------------
# Типы базисного вектора.
# enum ksBasisVectorTypeEnum
  ksTangentVector       = 0;
  ksNormalVector        = 1;
  ksBinormalVector      = 2;

#------------------------------------------------------------------------------
# Тип соединения кривых.
# enum ksConnectTypeEnum
  ksCTUnknown   = -1;
  ksCTPosition  = 0;
  ksCTTangent   = 1;
  ksCTNormal    = 2;
  ksCTSmooth    = 3;

#------------------------------------------------------------------------------
# Способ создания 3D дуги.
# enum ksArc3DBuildingTypeEnum
  ksArc3DByPoints       = 0;
  ksArc3DByCentre       = 1;
  ksArc3DByDirrection   = 2;
  ksArc3DByTanCurve     = 3;

#------------------------------------------------------------------------------
# Индекс параметра 3D дуги.
# enum ksArc3DParameterEnum
  ksArc3DCenter = 0;
  ksArc3DPoint1 = 1;
  ksArc3DPoint2 = 2;
  ksArc3DPoint3 = 3;
  ksArc3DAngle1 = 1;
  ksArc3DAngle2 = 2;
  ksArc3DRadius = 3;

#------------------------------------------------------------------------------
# Способ построения угла вращения.
# enum ksRotatedTypeEnum
  ksRTAngle     = 0;
  ksRTVertex    = 1;
  ksRTSurface   = 2;

#------------------------------------------------------------------------------
# Способ построения массива по сетке.
# enum ksLinearPatternBuildingTypeEnum
  ksLPSaveAll                   = 0;
  ksLPSaveAlongPerimeter        = 1;
  ksLPSaveAlongAxially          = 2;

#------------------------------------------------------------------------------
# Тип доступа к компоненту.
# enum ksPartAccessTypeEnum
  ksATUncertainty       = -1;
  ksATEditable          = 0;
  ksATReadOnly          = 1;
  ksATDisable           = 2;

#------------------------------------------------------------------------------
# Тип поверхности по сети точек.
# enum ksMeshPointsSurfaceBuildingTypeEnum
  ksMPByPoints  = 0;
  ksMPByPole    = 1;

#------------------------------------------------------------------------------
# Тип поверхности по облаку точек.
# enum ksCloudPointsSurfaceBuildingTypeEnum
  ksCLByPoints          = 0;
  ksCLByPole            = 1;
  ksCLPolyhedral        = 2;

#------------------------------------------------------------------------------
# Способ распознования сети точек.
# enum ksCloudTypeEnum
  ksCLAuto      = 0;
  ksCLLocalCS   = 1;
  ksCLScreen    = 2;

#------------------------------------------------------------------------------
# Способ фильтрации 3D объектов.
# enum ksObjectsFilter3DEnum
  ksFilterAll           = 0;
  ksFilterFaces         = 1;
  ksFilterEdges         = 2;
  ksFilterVertexs       = 3;
  ksFilterCPlanes       = 4;
  ksFilterCAxis         = 5;
  ksFilterParts         = 6;
  ksFilterBodies        = 7;
  ksFilterSurfaces      = 8;
  ksFilterSketches      = 9;
  ksFilterCurves        = 10;
  ksFilterCS            = 11;
  ksFilterControlPoints = 12;
  ksFilterPoints3D      = 13;
  ksFilterDesignations  = 14;
  ksFilterThread        = 15;

def RGB(r,g,b):
    return (b*256+g)*256+r






