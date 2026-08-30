const fs = require('fs');
const docx = require('docx');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType, VerticalAlign,
} = docx;

const FONT = 'Arial';
const ACCENT = '01696F';
const ACCENT_DARK = '0C4E54';
const GOLD = 'B98A1E';
const GOLD_DARK = '8A6712';
const MUTED = '7A7974';
const BORDER = 'D4D1CA';
const SURFACE = 'F7F6F2';
const TEXT = '28251D';

const rtlRun = (opts) => new TextRun({ font: FONT, rightToLeft: true, ...opts });

const P = (text, opts = {}) => new Paragraph({
  alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { after: 160, line: 300 },
  children: Array.isArray(text) ? text : [rtlRun({ text, size: 22, color: TEXT })], ...opts,
});

const H1 = (text, color = ACCENT_DARK, bar = ACCENT) => new Paragraph({
  heading: HeadingLevel.HEADING_1, alignment: AlignmentType.RIGHT, bidirectional: true,
  spacing: { before: 420, after: 200 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: bar, space: 6 } },
  children: [rtlRun({ text, size: 32, bold: true, color })],
});

const ActHeader = (text, fill = ACCENT_DARK) => new Paragraph({
  alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { before: 560, after: 260 },
  shading: { fill, type: ShadingType.CLEAR },
  children: [rtlRun({ text, size: 30, bold: true, color: 'FFFFFF' })],
});

const H3 = (text, color = TEXT) => new Paragraph({
  heading: HeadingLevel.HEADING_3, alignment: AlignmentType.RIGHT, bidirectional: true,
  spacing: { before: 220, after: 100 },
  children: [rtlRun({ text, size: 23, bold: true, color })],
});

const Meta = (label, value) => new Paragraph({
  alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { after: 60 },
  children: [
    rtlRun({ text: value, size: 20, color: TEXT }),
    rtlRun({ text: '  ', size: 20 }),
    rtlRun({ text: label + '  ', size: 20, bold: true, color: MUTED }),
  ],
});

const Stage = (text) => new Paragraph({
  alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { after: 140 },
  children: [rtlRun({ text: `[${text}]`, size: 20, italics: true, color: MUTED })],
});

const Narr = (text) => new Paragraph({
  alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { after: 160, line: 320 },
  children: [rtlRun({ text, size: 22, color: TEXT })],
});

const Line = (speaker, text) => new Paragraph({
  alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { after: 140, line: 320 },
  children: [
    rtlRun({ text: `${speaker}: `, size: 22, bold: true, color: ACCENT_DARK }),
    rtlRun({ text, size: 22, color: TEXT, italics: true }),
  ],
});

const Bullet = (t, color = ACCENT) => new Paragraph({
  alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { after: 80 },
  children: [rtlRun({ text: '•  ', size: 22, color }), rtlRun({ text: t, size: 22, color: TEXT })],
});

const Divider = () => new Paragraph({
  spacing: { before: 100, after: 260 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BORDER, space: 1 } },
  children: [new TextRun({ text: '' })],
});

function transition(text) {
  return new Paragraph({
    alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { before: 200, after: 260 },
    children: [rtlRun({ text: 'الانتقال: ', size: 20, bold: true, color: ACCENT_DARK }), rtlRun({ text, size: 20, color: MUTED })],
  });
}

// generic single-column "note box" table (grey/gold header)
function noteBox(title, lines, headerFill = ACCENT) {
  const cellMargins = { top: 100, bottom: 100, left: 150, right: 150 };
  const borderAll = { style: BorderStyle.SINGLE, size: 2, color: BORDER };
  const borders = { top: borderAll, bottom: borderAll, left: borderAll, right: borderAll, insideHorizontal: borderAll, insideVertical: borderAll };
  const tableWidth = 9360;
  const rows = [];
  rows.push(new TableRow({
    tableHeader: true,
    children: [new TableCell({
      width: { size: tableWidth, type: WidthType.DXA }, shading: { fill: headerFill, type: ShadingType.CLEAR },
      margins: cellMargins, verticalAlign: VerticalAlign.CENTER,
      children: [new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, children: [rtlRun({ text: title, size: 22, bold: true, color: 'FFFFFF' })] })],
    })],
  }));
  rows.push(new TableRow({
    children: [new TableCell({
      width: { size: tableWidth, type: WidthType.DXA }, shading: { fill: SURFACE, type: ShadingType.CLEAR }, margins: cellMargins,
      children: lines.map((l, i) => new Paragraph({
        alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { after: i === lines.length - 1 ? 0 : 90 },
        children: [rtlRun({ text: l.label ? `${l.label}: ` : '', size: 20, bold: true, color: MUTED }), rtlRun({ text: l.text, size: 20, color: TEXT })],
      })),
    })],
  }));
  return new Table({ width: { size: tableWidth, type: WidthType.DXA }, columnWidths: [tableWidth], borders, rows });
}

// ending card: condition header (gold) + variant lines table
function endingCard(code, title, condition, devFormula, variants) {
  const cellMargins = { top: 100, bottom: 100, left: 150, right: 150 };
  const borderAll = { style: BorderStyle.SINGLE, size: 2, color: BORDER };
  const borders = { top: borderAll, bottom: borderAll, left: borderAll, right: borderAll, insideHorizontal: borderAll, insideVertical: borderAll };
  const tableWidth = 9360;
  const rows = [];
  rows.push(new TableRow({
    tableHeader: true,
    children: [new TableCell({
      width: { size: tableWidth, type: WidthType.DXA }, shading: { fill: GOLD_DARK, type: ShadingType.CLEAR },
      margins: cellMargins, verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, children: [rtlRun({ text: `${code} — ${title}`, size: 24, bold: true, color: 'FFFFFF' })] }),
        new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { before: 60 }, children: [rtlRun({ text: 'شرط الترجيح: ', size: 20, bold: true, color: 'FFFFFF' }), rtlRun({ text: condition, size: 20, color: 'FFFFFF' })] }),
      ],
    })],
  }));
  rows.push(new TableRow({
    children: [new TableCell({
      width: { size: tableWidth, type: WidthType.DXA }, shading: { fill: 'FBF7EC', type: ShadingType.CLEAR }, margins: cellMargins,
      children: [new Paragraph({
        alignment: AlignmentType.RIGHT, bidirectional: true,
        children: [rtlRun({ text: 'ملاحظة للمبرمج — صيغة الفحص: ', size: 18, bold: true, color: MUTED }), rtlRun({ text: devFormula, size: 18, color: MUTED })],
      })],
    })],
  }));
  variants.forEach((v, idx) => {
    rows.push(new TableRow({
      children: [new TableCell({
        width: { size: tableWidth, type: WidthType.DXA }, shading: { fill: idx % 2 === 0 ? SURFACE : 'FBFBF9', type: ShadingType.CLEAR }, margins: cellMargins,
        children: [
          new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { after: 40 }, children: [rtlRun({ text: v.cond, size: 19, bold: true, color: GOLD_DARK })] }),
          new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, children: [rtlRun({ text: v.text, size: 20, italics: true, color: TEXT })] }),
        ],
      })],
    }));
  });
  return new Table({ width: { size: tableWidth, type: WidthType.DXA }, columnWidths: [tableWidth], borders, rows });
}

// trait bar row for the growth report table
function traitRow(label, code, hint, fill) {
  const cellMargins = { top: 80, bottom: 80, left: 150, right: 150 };
  const borderAll = { style: BorderStyle.SINGLE, size: 2, color: BORDER };
  return new TableRow({
    children: [
      new TableCell({
        width: { size: 2400, type: WidthType.DXA }, margins: cellMargins, shading: { fill: SURFACE, type: ShadingType.CLEAR },
        borders: { top: borderAll, bottom: borderAll, left: borderAll, right: borderAll },
        children: [new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, children: [rtlRun({ text: `${label} (${code})`, size: 20, bold: true, color: TEXT })] })],
      }),
      new TableCell({
        width: { size: 6960, type: WidthType.DXA }, margins: cellMargins, shading: { fill: 'FBFBF9', type: ShadingType.CLEAR },
        borders: { top: borderAll, bottom: borderAll, left: borderAll, right: borderAll },
        children: [new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, children: [rtlRun({ text: hint, size: 19, color: MUTED })] })],
      }),
    ],
  });
}

const children = [];

// ---------------- COVER ----------------
children.push(
  new Paragraph({ spacing: { before: 800 }, children: [] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, children: [rtlRun({ text: 'مَسَارُك', size: 64, bold: true, color: ACCENT_DARK })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { before: 160 }, children: [rtlRun({ text: 'الحوار الكامل للشاشة — الجزء الثالث', size: 34, bold: true, color: ACCENT })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { before: 80 }, children: [rtlRun({ text: 'محرك النهايات الأربع (E1–E4) وتقرير النمو السردي', size: 24, color: TEXT })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { before: 40 }, children: [rtlRun({ text: 'يكمل هذا الجزء الحوار الكامل من P0 حتى C12، ويغلق القصة الأساسية للنسخة MVP', size: 20, italics: true, color: MUTED })] }),
  new Paragraph({ spacing: { before: 600 }, children: [] }),
);

// ---------------- SECTION 0 ----------------
children.push(H1('0. منطق الفحص وقواعد الكتابة'));
children.push(Narr('يُفحص ترتيب النهايات بأولوية ثابتة: إذا تحقق شرط E1 تتوقف عملية الاختيار عنده. وإلا يُفحص شرط E2، ثم E3، وإلا تُطبَّق E4 كنتيجة افتراضية. لا تظهر نهاية خامسة أبدًا؛ التنويع يقتصر على جملة إضافية داخل النهاية نفسها حسب مستوى المال أو الضغط.'));

children.push(noteBox('صيغة الفحص الكاملة (شبه-كود للمبرمج)', [
  { label: '', text: 'إذا (سمعة R ≥ 60) و(اتزان B ≥ 55) و(انضباط D ≥ 55) وعلم «مسؤولية» موجود، وأحد الأعلام {استدامة، تجربة_محدودة، إعادة_هيكلة} موجود ← E1' },
  { label: '', text: 'وإلا إذا (جرأة C ≥ 65) و(مال M ≥ 55)، وأحد الأعلام {توسع_سريع، شراكة_سريعة، دين} موجود ← E2' },
  { label: '', text: 'وإلا إذا (معرفة K ≥ 60) و(انضباط D ≥ 55) ← E3' },
  { label: '', text: 'وإلا ← E4 (تشمل أيضًا: سمعة R منخفضة مع ضغط P مرتفع، أو تكرار علم «خروج» بلا خطة واضحة)' },
], ACCENT_DARK));

children.push(new Paragraph({ spacing: { before: 200 }, children: [] }));
children.push(H3('قواعد إضافية مطبَّقة في هذا الجزء'));
children.push(Bullet('كل نهاية مشهد سينمائي قصير، لا شاشة نص جافة؛ تُختم بجملة عنوان (tagline) واحدة.'));
children.push(Bullet('الجمل الإضافية حسب المال أو الضغط تُضاف داخل المشهد نفسه، ولا تُنشئ فرعًا جديدًا أو نهاية خامسة.'));
children.push(Bullet('لا تُستخدم لغة تشخيص نفسي، ولا يُربط المال أو نوع الشهادة بقيمة اللاعب، حتى في E4.'));
children.push(Bullet('الأرقام الدقيقة للسمات لا تظهر خلال القصة، لكنها تظهر بوضوح في تقرير النمو السردي الختامي فقط.'));

children.push(Divider());

// ---------------- EVAL node ----------------
children.push(H1('عقدة التقييم — EVAL'));
children.push(Meta('المكان والزمان', 'انتقال بصري بعد إغلاق باب المصعد في نهاية C12'));
children.push(Stage('الشاشة تعتم للحظة قصيرة (لا يزيد عن ثانيتين)، بلا أي نص ظاهر للاعب'));
children.push(Narr('في هذه اللحظة يطبّق المحرك تراكم الأثر النهائي على جميع السمات، ثم يحصر القيم الخمس بين 0 و100، ثم يفحص شروط النهايات بالترتيب أعلاه.'));
children.push(transition('عند تحديد النهاية، تنتقل الشاشة إلى مشهد V3 الخاص بها، ثم إلى تقرير النمو السردي.'));
children.push(Divider());

// ---------------- E1 ----------------
children.push(H1('E1 — صانع الأثر', GOLD_DARK, GOLD));
children.push(Meta('المكان والزمان', 'مكتب هادئ، نهاية يوم عمل، إضاءة دافئة'));
children.push(Meta('الشخصيات', 'نوف، ومتدرب جديد في الخلفية'));

children.push(Stage('تقرير مفتوح على الشاشة يحمل ختم «مكتمل ضمن الموعد»'));
children.push(Narr('تنمو المبادرة أو الإدارة على مراحل، لا بضجيج إعلان كبير، بل بتراكم قرارات صغيرة صحيحة. الفريق الذي بنيته لم يعد يحتاج إشرافك على كل تفصيل.'));
children.push(Stage('نوف، التي كانت يومًا الأحدث في الفريق، تشرح لمتدرب جديد قاعدة سمعتها منك يومًا ما'));
children.push(Line('نوف', 'القاعدة بسيطة: قل الحقيقة أولًا، ثم أصلح الخطأ.'));
children.push(Narr('يغلق اللاعب تقرير اليوم في موعده، دون تأخير ودون إنذار أخير مزعج. الشاشة تتلاشى إلى الأسود بهدوء.'));

children.push(endingCard('E1', 'صانع الأثر',
  'سمعة ≥ 60، واتزان ≥ 55، وانضباط ≥ 55، مع علم «مسؤولية»، وأحد الأعلام: استدامة / تجربة_محدودة / إعادة_هيكلة',
  'R≥60 ∧ B≥55 ∧ D≥55 ∧ flag(مسؤولية) ∧ (flag(استدامة) ∨ flag(تجربة_محدودة) ∨ flag(إعادة_هيكلة))',
  [
    { cond: 'إذا كان الضغط P ≥ 65', text: 'النجاح هنا حقيقي، لكن جدولك ما زال يحتاج مساحة بيضاء لم تمنحها لنفسك بعد.' },
    { cond: 'إذا كان الضغط P < 65', text: 'النجاح هنا يأتي مصحوبًا براحة نادرًا ما شعرت بها من قبل.' },
    { cond: 'إذا كان المال M ≥ 70', text: 'الاستقرار المالي لم يكن الهدف، لكنه جاء نتيجة طبيعية للثقة التي بنيتها.' },
    { cond: 'إذا كان المال M < 40', text: 'الأرقام في حسابك متواضعة، لكن ما بنيته من ثقة لا يُقاس بها.' },
  ],
));
children.push(P([rtlRun({ text: 'جملة العنوان: ', size: 20, bold: true, color: MUTED }), rtlRun({ text: '«لم تخترْ بين النجاح والفشل. اخترتَ أن يبقى أثرك بعدك.»', size: 22, italics: true, color: TEXT })], { spacing: { before: 200, after: 200 } }));
children.push(Divider());

// ---------------- E2 ----------------
children.push(H1('E2 — الرائد الجريء', GOLD_DARK, GOLD));
children.push(Meta('المكان والزمان', 'مكتب جديد في مدينة أخرى، صناديق لم تُفرغ بعد'));
children.push(Meta('الشخصيات', 'سارة (في نسخة الضغط المرتفع فقط)'));

children.push(Stage('شاشة تعرض خريطة بها عدة نقاط مضيئة تمثل مدنًا جديدة'));
children.push(Narr('تصل الفكرة إلى مدن أكثر بسرعة تفوق ما تخيلته في المصعد الزجاجي. الاسم الذي بدأ على طاولة صغيرة أصبح يُذكر في اجتماعات لا تحضرها بنفسك.'));

children.push(H3('نسخة المشهد الختامي حسب مستوى الضغط'));
children.push(Stage('إذا كان الضغط P ≥ 65 — هاتفك يعرض تقويمًا بلا أي مساحة فارغة حتى نهاية الشهر'));
children.push(Narr('تنتهي اللقطة بجدول مزدحم، ورسالة من سارة تصلك في وقت متأخر.'));
children.push(Line('سارة', 'وسّعنا كل شيء إلا وقتك. رتّب نجاحك قبل أن يرتبك هو.'));
children.push(Stage('إذا كان الضغط P < 65 — مكتب مرتب، لحظة هدوء نادرة بين اجتماعين'));
children.push(Narr('تظهر قيادة حاسمة مع مساحة كافية للتنفس؛ القرارات السريعة التي اتخذتها لم تُكلفك نفسك.'));

children.push(endingCard('E2', 'الرائد الجريء',
  'جرأة ≥ 65 ومال ≥ 55، مع أحد الأعلام: توسع_سريع / شراكة_سريعة / دين',
  'C≥65 ∧ M≥55 ∧ (flag(توسع_سريع) ∨ flag(شراكة_سريعة) ∨ flag(دين))',
  [
    { cond: 'إذا كان الضغط P ≥ 65', text: 'المشهد الختامي = نسخة «جدول مزدحم» مع رسالة سارة أعلاه.' },
    { cond: 'إذا كان الضغط P < 65', text: 'المشهد الختامي = نسخة «قيادة حاسمة مع مساحة للتنفس» أعلاه.' },
  ],
));
children.push(P([rtlRun({ text: 'جملة العنوان: ', size: 20, bold: true, color: MUTED }), rtlRun({ text: '«التقطت الفرصة قبل أن تفوتك. الثمن كان السرعة نفسها.»', size: 22, italics: true, color: TEXT })], { spacing: { before: 200, after: 200 } }));
children.push(Divider());

// ---------------- E3 ----------------
children.push(H1('E3 — الخبير الموثوق', GOLD_DARK, GOLD));
children.push(Meta('المكان والزمان', 'قاعة تدريب أو مختبر صغير، لوح مليء بالملاحظات'));
children.push(Meta('الشخصيات', 'زميل يتسلم الإدارة اليومية، وحضور ينتظرون الاستشارة'));

children.push(Narr('تتخصص الشخصية وتصبح مرجعًا يطلبه الآخرون للمسائل الصعبة، لا لأنها الأعلى صوتًا، بل لأنها الأكثر دقة في وقت الحاجة.'));
children.push(Line('أحد الحضور', 'قالوا لي اسأل هنا قبل أي مكان آخر.'));
children.push(Stage('تسلّم ملف الإدارة اليومية لزميل أثبت جاهزيته، وتفتح مكانه ملفًا جديدًا على مكتبك'));
children.push(Narr('تسلّم الإدارة لمن هو أنسب، وتفتح ملف تعلم جديد، كأن القصة لم تنتهِ بل بدّلت اتجاهها فقط.'));

children.push(endingCard('E3', 'الخبير الموثوق',
  'معرفة ≥ 60 وانضباط ≥ 55، ولم تتحقق شروط E1 أو E2',
  'K≥60 ∧ D≥55 ∧ ¬(شرط E1) ∧ ¬(شرط E2)',
  [
    { cond: 'إذا كان المال M ≥ 70', text: 'الخبرة هنا رفعت دخلك أيضًا، لكنها لم تكن الهدف من الرحلة.' },
    { cond: 'إذا كان المال M < 40', text: 'الحساب البنكي لا يعكس كل ما راكمته من معرفة يطلبها الآخرون.' },
    { cond: 'إذا كان الضغط P ≥ 65', text: 'التركيز العميق أنهكك أحيانًا، لكنه أعطاك دقة نادرة.' },
    { cond: 'إذا كان الضغط P < 65', text: 'التخصص منحك سرعة هادئة في اتخاذ القرار، لا توتر السباق.' },
  ],
));
children.push(P([rtlRun({ text: 'جملة العنوان: ', size: 20, bold: true, color: MUTED }), rtlRun({ text: '«لم تُطارد كل فرصة. طاردتَ فهمًا أعمق لفرصة واحدة.»', size: 22, italics: true, color: TEXT })], { spacing: { before: 200, after: 200 } }));
children.push(Divider());

// ---------------- E4 ----------------
children.push(H1('E4 — البداية الثانية', GOLD_DARK, GOLD));
children.push(Meta('المكان والزمان', 'إضاءة صباحية عادية، طاولة صغيرة في المنزل'));
children.push(Meta('الشخصيات', 'الأستاذ فهد (رسالة قصيرة)'));

children.push(Stage('لا شاشة فشل ولا موسيقى حزينة؛ المشهد هادئ وعادي كأي صباح آخر'));
children.push(Narr('تعود الشخصية إلى دفتر C1 القديم، ذلك الذي أهداه الأستاذ فهد يومًا، وتفتح صفحة جديدة فيه.'));
children.push(Narr('تكتب ما تعلمته بخط واضح، ثم تسدد أول التزام صغير بقي عليك، وتتخذ خطوة أصغر لكن أوضح من أي خطوة سابقة.'));
children.push(Stage('رسالة قصيرة تصل من الأستاذ فهد'));
children.push(Line('الأستاذ فهد', 'البداية الثانية ليست فشل الأولى؛ إنها المعلومة التي جمعتها منها.'));

children.push(endingCard('E4', 'البداية الثانية',
  'النتيجة الافتراضية عند عدم تحقق أي شرط أعلى، أو سمعة منخفضة مع ضغط مرتفع، أو تكرار علم «خروج» بلا خطة واضحة',
  'default ∨ (R<40 ∧ P≥65) ∨ (count(flag(خروج)) ≥ 2 ∧ ¬flag(خطة_واضحة))',
  [
    { cond: 'إذا كانت السمعة R < 40 والضغط P ≥ 65', text: 'أثقل ما تحمله الآن ليس الدين، بل الحاجة لاستعادة ثقة من حولك، خطوة بخطوة.' },
    { cond: 'إذا كان السبب تكرار علم «خروج» بلا خطة', text: 'هذه ليست أول مرة تبدأ فيها من جديد، لكنها أول مرة تكتب فيها خطة قبل الخطوة.' },
  ],
));
children.push(P([rtlRun({ text: 'جملة العنوان: ', size: 20, bold: true, color: MUTED }), rtlRun({ text: '«لا نهاية لك هنا. فقط صفحة جديدة في الدفتر نفسه.»', size: 22, italics: true, color: TEXT })], { spacing: { before: 200, after: 100 } }));
children.push(transition('يفتح طور إعادة اللعب مع تلميحات جديدة تظهر في الشاشة التالية، دون اعتبار هذه النهاية عقوبة.'));
children.push(Divider());

// ---------------- GROWTH REPORT ----------------
children.push(ActHeader('تقرير النمو السردي — الشاشة الختامية V3', GOLD_DARK));

children.push(Meta('المكان والزمان', 'شاشة تقرير نهائية، بعد مشهد النهاية مباشرة'));
children.push(Stage('تنبيه ثابت يظهر أعلى كل تقرير، بخط أصغر ولون محايد'));
children.push(Narr('هذا المؤشر سردي وترفيهي، وليس اختبارًا نفسيًا أو حكمًا على نضجك في الحياة الواقعية.'));

children.push(H3('حساب مؤشر النمو'));
children.push(P([
  rtlRun({ text: 'مؤشر النمو السردي = ', size: 22, bold: true, color: ACCENT_DARK }),
  rtlRun({ text: '(المعرفة K + الانضباط D + الجرأة C + الاتزان B + السمعة R) ÷ 5', size: 22, color: TEXT }),
]));
children.push(Narr('لا يدخل المال M ولا الضغط P في هذا المتوسط؛ الثراء ليس نضجًا، والضغط ليس نقصًا أخلاقيًا. يظهر المال والضغط في التقرير كسياق منفصل فقط.'));

children.push(H3('عرض السمات الخمس على الشاشة'));
{
  const borderAll = { style: BorderStyle.SINGLE, size: 2, color: BORDER };
  const rows = [
    new TableRow({ tableHeader: true, children: [
      new TableCell({ width: { size: 2400, type: WidthType.DXA }, shading: { fill: ACCENT, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 150, right: 150 }, borders: { top: borderAll, bottom: borderAll, left: borderAll, right: borderAll }, children: [new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, children: [rtlRun({ text: 'السمة', size: 20, bold: true, color: 'FFFFFF' })] })] }),
      new TableCell({ width: { size: 6960, type: WidthType.DXA }, shading: { fill: ACCENT, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 150, right: 150 }, borders: { top: borderAll, bottom: borderAll, left: borderAll, right: borderAll }, children: [new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, children: [rtlRun({ text: 'كيف تُعرض في التقرير', size: 20, bold: true, color: 'FFFFFF' })] })] }),
    ]}),
    traitRow('المعرفة', 'K', 'شريط تقدم من 0 إلى 100 مع رقم دقيق — يظهر هنا فقط، لا خلال القصة.'),
    traitRow('الانضباط', 'D', 'شريط تقدم من 0 إلى 100 مع رقم دقيق.'),
    traitRow('الجرأة', 'C', 'شريط تقدم من 0 إلى 100 مع رقم دقيق.'),
    traitRow('الاتزان', 'B', 'شريط تقدم من 0 إلى 100 مع رقم دقيق.'),
    traitRow('السمعة', 'R', 'شريط تقدم من 0 إلى 100 مع رقم دقيق.'),
    traitRow('المال', 'M', 'رقم سياقي منفصل أسفل الأشرطة الخمسة، بعنوان «سياق القصة» لا «نتيجة».'),
    traitRow('الضغط', 'P', 'رقم سياقي منفصل، يُستخدم فقط لاختيار صياغة الخاتمة أعلاه.'),
  ];
  children.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2400, 6960], borders: { top: borderAll, bottom: borderAll, left: borderAll, right: borderAll, insideHorizontal: borderAll, insideVertical: borderAll }, rows }));
}

children.push(new Paragraph({ spacing: { before: 260 }, children: [] }));
children.push(H3('نطاقات المتوسط ولغة التقرير'));
{
  const borderAll = { style: BorderStyle.SINGLE, size: 2, color: BORDER };
  const cellMargins = { top: 90, bottom: 90, left: 150, right: 150 };
  const bands = [
    ['0–39', 'بداية واعية', 'تعرفت إلى أنماط قراراتك، والخطوة التالية أصغر وأكثر وضوحًا.'],
    ['40–59', 'نمو واضح', 'تملك نقاط قوة متكررة، لكن بعض القرارات لا تزال قصيرة الأفق.'],
    ['60–79', 'نضج عملي', 'تربط الطموح بالتعلم والمسؤولية وتتعافى من التعثر.'],
    ['80–100', 'أثر راسخ', 'تبني قرارات قابلة للاستمرار وتمنح من حولك مساحة للنمو.'],
  ];
  const headerRow = new TableRow({ tableHeader: true, children: [
    new TableCell({ width: { size: 1600, type: WidthType.DXA }, shading: { fill: ACCENT, type: ShadingType.CLEAR }, margins: cellMargins, borders: { top: borderAll, bottom: borderAll, left: borderAll, right: borderAll }, children: [new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, children: [rtlRun({ text: 'المتوسط', size: 19, bold: true, color: 'FFFFFF' })] })] }),
    new TableCell({ width: { size: 2400, type: WidthType.DXA }, shading: { fill: ACCENT, type: ShadingType.CLEAR }, margins: cellMargins, borders: { top: borderAll, bottom: borderAll, left: borderAll, right: borderAll }, children: [new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, children: [rtlRun({ text: 'الوصف الظاهر', size: 19, bold: true, color: 'FFFFFF' })] })] }),
    new TableCell({ width: { size: 5360, type: WidthType.DXA }, shading: { fill: ACCENT, type: ShadingType.CLEAR }, margins: cellMargins, borders: { top: borderAll, bottom: borderAll, left: borderAll, right: borderAll }, children: [new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, children: [rtlRun({ text: 'لغة التقرير', size: 19, bold: true, color: 'FFFFFF' })] })] }),
  ]});
  const rows = [headerRow].concat(bands.map(([range, label, text], i) => new TableRow({ children: [
    new TableCell({ width: { size: 1600, type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? SURFACE : 'FBFBF9', type: ShadingType.CLEAR }, margins: cellMargins, borders: { top: borderAll, bottom: borderAll, left: borderAll, right: borderAll }, children: [new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, children: [rtlRun({ text: range, size: 19, bold: true, color: TEXT })] })] }),
    new TableCell({ width: { size: 2400, type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? SURFACE : 'FBFBF9', type: ShadingType.CLEAR }, margins: cellMargins, borders: { top: borderAll, bottom: borderAll, left: borderAll, right: borderAll }, children: [new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, children: [rtlRun({ text: label, size: 19, bold: true, color: ACCENT_DARK })] })] }),
    new TableCell({ width: { size: 5360, type: WidthType.DXA }, shading: { fill: i % 2 === 0 ? SURFACE : 'FBFBF9', type: ShadingType.CLEAR }, margins: cellMargins, borders: { top: borderAll, bottom: borderAll, left: borderAll, right: borderAll }, children: [new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, children: [rtlRun({ text, size: 19, italics: true, color: TEXT })] })] }),
  ]})));
  children.push(new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [1600, 2400, 5360], borders: { top: borderAll, bottom: borderAll, left: borderAll, right: borderAll, insideHorizontal: borderAll, insideVertical: borderAll }, rows }));
}

children.push(new Paragraph({ spacing: { before: 260 }, children: [] }));
children.push(H3('عناصر إضافية في شاشة التقرير'));
children.push(Bullet('ملخص الأعلام السردية الأساسية (studyPath، مرشد، شفافية، صندوق_طوارئ، مسؤولية...) بعنوان «القرارات التي شكّلت قصتك» بصياغة نصية لا رقمية.'));
children.push(Bullet('زر واحد: «العب مسارًا آخر»، يعيد اللاعب إلى P0 مع الاحتفاظ بسجل النهايات التي وصل إليها سابقًا.'));
children.push(Bullet('لا تُستخدم أي عبارة تقييمية عن الذكاء أو النجاح الشخصي خارج سياق اللعبة؛ التقرير يصف القصة لا اللاعب.'));

children.push(Divider());

// ---------------- CLOSING ----------------
children.push(H1('اكتمال القصة الأساسية'));
children.push(P([rtlRun({ text: 'بهذا الجزء تكتمل «القصة الأساسية وشجرة القرارات» بالكامل: ', size: 22 }), rtlRun({ text: 'المقدمة P0، الفصول C1–C12، ومحرك النهايات E1–E4 مع تقرير النمو السردي.', size: 22, bold: true })]));
children.push(P('حسب معيار اكتمال MVP في حزمة التصميم، أصبحت كل السلسلة قابلة للتحويل إلى بيانات لعبة (scenes.json) واختبار آلي للمسارات الأربعة الكاملة من البداية حتى كل نهاية.'));
children.push(P('الخطوات التالية المقترحة: بناء ملفات scenes.json لبقية الفصول (بعد أن أُنجزت P0–C4 كدفعة أولى)، ثم اختبار مسارات آلية تصل إلى كل نهاية من E1 إلى E4 للتأكد من خلو الشجرة من عقد مفقودة.'));

const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: 22 }, paragraph: { bidirectional: true, alignment: AlignmentType.RIGHT } } },
  },
  sections: [
    {
      properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } } },
      children,
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync('/home/user/workspace/masarak_dialogue_endings.docx', buf);
  console.log('done');
});
