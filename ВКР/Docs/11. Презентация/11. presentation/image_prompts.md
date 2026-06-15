# Промпты для генерации изображений

Два плейсхолдера в презентации:
- **Слайд 2** (Актуальность) → `images/intro.png`
- **Слайд 11** (Рекомендации по внедрению) → `images/conveyor.png`

Когда картинки сгенерированы, положить их по этим путям и заменить `<div class="img-placeholder">...</div>` в `index.html` на `<img src="images/intro.png">` и `<img src="images/conveyor.png">` соответственно.

Финальное соотношение сторон — примерно **3:2** (горизонтальное), разрешение от **1280×854** и выше. Формат PNG/JPEG без альфа-канала.

---

## Промпт 1 — `intro.png` (слайд «Актуальность»)

**Цель:** показать проблему — бракованный картон на промышленной линии, подчеркнуть что ручной контроль не справляется.

### Английский вариант (для Midjourney / DALL·E / Stable Diffusion / Flux)

```
Industrial corrugated cardboard production line, close-up of a wide conveyor belt carrying flat cardboard sheets moving left-to-right, one sheet in the foreground clearly has visible surface defects — a long scratch and a small dent — while other sheets behind it look clean. Cool factory lighting, metallic machinery and rollers in the background, slightly blurred, realistic industrial photography style, no people, no text, no logos. Composition has strong leading lines of the conveyor drawing the eye to the defective sheet. Slight shallow depth of field. Color palette: neutral grays, warm beige of cardboard, cold blue industrial light. Photorealistic, documentary style, 35mm lens, f/5.6, high detail. Aspect ratio 3:2.
```

### Русский вариант (перевод / уточнение для себя)

Промышленная линия производства гофрокартона, крупный план широкой ленты конвейера с плоскими листами картона, движущимися слева направо. На одном листе на переднем плане чётко видны дефекты — длинная царапина и небольшая вмятина; остальные листы позади чистые. Фабричное холодное освещение, металлические механизмы и ролики на заднем фоне, лёгкое размытие. Стиль — документальная промышленная фотография. Без людей, без текста, без логотипов. Композиция: конвейер уходит в глубину и ведёт взгляд к бракованному листу. Неглубокая глубина резкости. Палитра: нейтральный серый, тёплый бежевый картона, холодный синий индустриальный свет. 3:2.

### Настройки (если Midjourney)

```
--ar 3:2 --style raw --s 100 --q 2
```

### Альтернативная формулировка (если хочется более «плакатный» стиль)

```
Editorial-style industrial illustration: a cardboard factory conveyor with stacked flat brown sheets, one sheet highlighted in soft red glow to indicate a defect, clean modern vector-like composition mixed with photographic realism, minimalist color palette dominated by cardboard beige and factory blue-grey, no text. Aspect ratio 3:2.
```

---

## Промпт 2 — `conveyor.png` (слайд «Рекомендации по внедрению»)

**Цель:** схематично показать линию контроля с камерами, Jetson Orin и пневматическим отбраковщиком — рекомендуемый этап 1.

### Английский вариант

```
Technical isometric illustration of a cardboard quality-control inspection station on a factory conveyor. Show: (1) a wide conveyor belt carrying flat brown cardboard sheets moving left to right; (2) two industrial cameras mounted on an overhead frame above the belt, with visible LED strip lights flanking them; (3) an edge AI computer box (NVIDIA Jetson style, compact black enclosure with small green status LED) mounted on a nearby control cabinet; (4) a pneumatic rejector arm (air jet / piston) positioned downstream of the cameras, aimed diagonally at the belt, about to push one defective sheet off into a small reject bin; (5) a PLC cabinet on the side with a few cable runs to camera and rejector. Labels in small clean sans-serif: "camera", "LED light", "Jetson Orin", "PLC", "reject". Clean technical diagram style, isometric perspective, flat modern vector illustration, palette: industrial blue (#00549F), warm cardboard beige, neutral gray, accent yellow-green for the reject action. White background. No photorealism — crisp flat vectors with minimal shading. Aspect ratio 3:2.
```

### Русский вариант

Технический изометрический рисунок станции контроля качества картона на заводском конвейере. Элементы:
1. Широкий ленточный конвейер с плоскими бежевыми листами картона, движущимися слева направо.
2. Две промышленные камеры на надконвейерной раме, по бокам — LED-полосы подсветки.
3. Компактный edge-AI компьютер в стиле NVIDIA Jetson (чёрный корпус, зелёный статусный светодиод) на боковом шкафу.
4. Пневматический отбраковщик (струя воздуха / поршень) ниже по потоку, наклонно направленный на ленту, сбрасывает один бракованный лист в отдельный лоток.
5. Шкаф PLC сбоку с несколькими кабелями к камерам и отбраковщику.
6. Мелкие подписи на изображении: «camera», «LED light», «Jetson Orin», «PLC», «reject».

Стиль: плоская векторная техническая иллюстрация, изометрия, минимум теней. Палитра: индустриальный синий КФУ (#00549F), тёплый бежевый картона, нейтральный серый, акцент жёлто-зелёным в момент отбраковки. Белый фон. Без фотореализма. Соотношение 3:2.

### Настройки (если Midjourney)

```
--ar 3:2 --style raw --s 250
```

### Альтернатива (если хочется фото, а не схему)

```
Photorealistic industrial photography of an inspection section on a cardboard conveyor: two overhead smart-cameras with ring lights, pneumatic rejector arm sweeping a defective sheet into a side bin, small Jetson-style edge device visible on the side cabinet, clean factory environment, cool neutral lighting. No people, no text, no logos. Aspect ratio 3:2.
```

---

## После генерации

1. Сохранить файлы как `images/intro.png` и `images/conveyor.png`.
2. В `index.html` найти два блока `<div class="img-placeholder">...</div>` и заменить на:

```html
<!-- слайд 2 -->
<img src="images/intro.png" alt="Брак в производстве картона" style="flex: 1; min-height: 0;">

<!-- слайд 11 -->
<img src="images/conveyor.png" alt="Схема станции контроля" style="width: 100%; margin-top: 10px; border-radius: 4px;">
```

3. Проверить что картинки не разламывают layout — `img` в slide-body уже имеет `max-width: 100%; max-height: 100%; object-fit: contain;` в style.css.
