const fs = require('fs');
const docx = require('docx');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType, VerticalAlign,
} = docx;

const FONT = 'Arial';
const ACCENT = '01696F';
const ACCENT_DARK = '0C4E54';
const MUTED = '7A7974';
const BORDER = 'D4D1CA';
const SURFACE = 'F7F6F2';
const TEXT = '28251D';

const rtlRun = (opts) => new TextRun({ font: FONT, rightToLeft: true, ...opts });

const P = (text, opts = {}) => new Paragraph({
  alignment: AlignmentType.RIGHT,
  bidirectional: true,
  spacing: { after: 160, line: 300 },
  children: Array.isArray(text) ? text : [rtlRun({ text, size: 22, color: TEXT })],
  ...opts,
});

const H1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1,
  alignment: AlignmentType.RIGHT,
  bidirectional: true,
  spacing: { before: 420, after: 200 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: ACCENT, space: 6 } },
  children: [rtlRun({ text, size: 32, bold: true, color: ACCENT_DARK })],
});

const H2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2,
  alignment: AlignmentType.RIGHT,
  bidirectional: true,
  spacing: { before: 300, after: 140 },
  children: [rtlRun({ text, size: 26, bold: true, color: ACCENT })],
});

const H3 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_3,
  alignment: AlignmentType.RIGHT,
  bidirectional: true,
  spacing: { before: 220, after: 100 },
  children: [rtlRun({ text, size: 23, bold: true, color: TEXT })],
});

const Meta = (label, value) => new Paragraph({
  alignment: AlignmentType.RIGHT,
  bidirectional: true,
  spacing: { after: 60 },
  children: [
    rtlRun({ text: value, size: 20, color: TEXT }),
    rtlRun({ text: '  ', size: 20 }),
    rtlRun({ text: label + '  ', size: 20, bold: true, color: MUTED }),
  ],
});

const Stage = (text) => new Paragraph({
  alignment: AlignmentType.RIGHT,
  bidirectional: true,
  spacing: { after: 140 },
  children: [rtlRun({ text: `[${text}]`, size: 20, italics: true, color: MUTED })],
});

const Narr = (text) => new Paragraph({
  alignment: AlignmentType.RIGHT,
  bidirectional: true,
  spacing: { after: 160, line: 320 },
  indent: { left: 0 },
  children: [rtlRun({ text, size: 22, color: TEXT })],
});

const Line = (speaker, text) => new Paragraph({
  alignment: AlignmentType.RIGHT,
  bidirectional: true,
  spacing: { after: 140, line: 320 },
  children: [
    rtlRun({ text: `${speaker}: `, size: 22, bold: true, color: ACCENT_DARK }),
    rtlRun({ text, size: 22, color: TEXT, italics: true }),
  ],
});

const Divider = () => new Paragraph({
  spacing: { before: 100, after: 260 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BORDER, space: 1 } },
  children: [new TextRun({ text: '' })],
});

// ---- Decision card as a table ----
function decisionCard(promptTitle, promptText, options) {
  const cellMargins = { top: 100, bottom: 100, left: 150, right: 150 };
  const borderAll = { style: BorderStyle.SINGLE, size: 2, color: BORDER };
  const borders = { top: borderAll, bottom: borderAll, left: borderAll, right: borderAll, insideHorizontal: borderAll, insideVertical: borderAll };
  const tableWidth = 9360;

  const rows = [];
  rows.push(new TableRow({
    tableHeader: true,
    children: [new TableCell({
      columnSpan: 1,
      width: { size: tableWidth, type: WidthType.DXA },
      shading: { fill: ACCENT, type: ShadingType.CLEAR },
      margins: cellMargins,
      verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, children: [rtlRun({ text: `نقطة القرار — ${promptTitle}`, size: 22, bold: true, color: 'FFFFFF' })] }),
        new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { before: 60 }, children: [rtlRun({ text: promptText, size: 20, color: 'FFFFFF' })] }),
      ],
    })],
  }));

  options.forEach((opt, idx) => {
    rows.push(new TableRow({
      children: [new TableCell({
        width: { size: tableWidth, type: WidthType.DXA },
        shading: { fill: idx % 2 === 0 ? SURFACE : 'FBFBF9', type: ShadingType.CLEAR },
        margins: cellMargins,
        children: [
          new Paragraph({
            alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { after: 60 },
            children: [
              rtlRun({ text: `  ⟨${opt.code}⟩ `, size: 18, color: MUTED }),
              rtlRun({ text: opt.button, size: 22, bold: true, color: TEXT }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { after: 40 },
            children: [
              rtlRun({ text: 'استجابة فورية: ', size: 20, bold: true, color: ACCENT_DARK }),
              rtlRun({ text: opt.response, size: 20, italics: true, color: TEXT }),
              rtlRun({ text: '  ' + opt.tags, size: 20, color: ACCENT }),
            ],
          }),
          new Paragraph({
            alignment: AlignmentType.RIGHT, bidirectional: true,
            children: [
              rtlRun({ text: 'ملاحظة للمبرمج (لا تظهر للاعب): ', size: 18, bold: true, color: MUTED }),
              rtlRun({ text: opt.dev, size: 18, color: MUTED }),
            ],
          }),
        ],
      })],
    }));
  });

  return new Table({
    width: { size: tableWidth, type: WidthType.DXA },
    columnWidths: [tableWidth],
    borders,
    rows,
  });
}

function transition(text) {
  return new Paragraph({
    alignment: AlignmentType.RIGHT,
    bidirectional: true,
    spacing: { before: 200, after: 260 },
    children: [
      rtlRun({ text: 'الانتقال: ', size: 20, bold: true, color: ACCENT_DARK }),
      rtlRun({ text, size: 20, color: MUTED }),
    ],
  });
}

const children = [];

// ---------------- COVER ----------------
children.push(
  new Paragraph({ spacing: { before: 800 }, children: [] }),
  new Paragraph({
    alignment: AlignmentType.RIGHT, bidirectional: true,
    children: [rtlRun({ text: 'مَسَارُك', size: 64, bold: true, color: ACCENT_DARK })],
  }),
  new Paragraph({
    alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { before: 160 },
    children: [rtlRun({ text: 'الحوار الكامل للشاشة', size: 34, bold: true, color: ACCENT })],
  }),
  new Paragraph({
    alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { before: 80 },
    children: [rtlRun({ text: 'المقدمة القابلة للعب (P0) والفصول من الأول إلى الرابع (C1–C4)', size: 24, color: TEXT })],
  }),
  new Paragraph({
    alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { before: 40 },
    children: [rtlRun({ text: 'نسخة جاهزة للتنفيذ — الدفعة البرمجية الأولى، مبنية على حزمة التصميم السردي', size: 20, italics: true, color: MUTED })],
  }),
  new Paragraph({ spacing: { before: 600 }, children: [] }),
);

// ---------------- SECTION 0: PRODUCTION NOTES ----------------
children.push(H1('0. ملاحظات القراءة والإنتاج'));
children.push(Narr('يغطي هذا المستند الدفعة الأولى المقترحة في حزمة التصميم: المقدمة وفصول «آخر جرس»، «ثلاثة أبواب»، «أول مبلغ يخصك»، و«العلامة الحمراء». الهدف اختبار اختيار الشخصية والمسار والسمات والحفظ والتطور البصري قبل كتابة بقية الفصول (C5–C12) ونهايات E1–E4، والتي تُكتب في دفعة لاحقة عند الطلب.'));

children.push(H3('بنية كل بطاقة مشهد'));
children.push(P([
  rtlRun({ text: 'المكان والزمان ← السرد (نص يظهر على الشاشة على شكل بطاقات قصيرة) ← الحوار المنسوب لشخصية باسمها ← نقطة القرار. عند نقطة القرار يرى اللاعب فقط نص الزر والاستجابة الفورية بعده؛ لا تظهر أي أرقام ولا رموز سمات على الشاشة.', size: 22 }),
]));

children.push(H3('دليل رموز السمات (للفريق فقط، لا تظهر للاعب)'));
children.push(P('K = المعرفة  ·  D = الانضباط  ·  C = الجرأة  ·  B = الاتزان  ·  R = السمعة  ·  M = المال  ·  P = الضغط'));

children.push(H3('قواعد الكتابة المطبَّقة في هذا النص'));
[
  'صيغة مخاطب محايدة، مضارع مستمر، وتجنّب لواحق التذكير والتأنيث حيث لا يلزم.',
  'نص كل زر بين 4 و10 كلمات، يبدأ بفعل، ولا يكشف الأثر الرقمي.',
  'بعد كل اختيار: جملة واحدة كاستجابة فورية، بلا حكم أخلاقي وبلا تشخيص.',
  'لا شعارات أو منصات مالية حقيقية؛ الأرقام أدوات سردية فقط وليست نصيحة مالية.',
  'مسارات C2 الثلاثة (جامعة/دبلوم/عمل) تُكتب كطبقة نص شرطي قصيرة فوق بنية مشتركة، لا كثلاث قصص منفصلة.',
].forEach(t => children.push(new Paragraph({
  alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { after: 80 },
  children: [rtlRun({ text: '•  ', size: 22, color: ACCENT }), rtlRun({ text: t, size: 22, color: TEXT })],
})));

children.push(Divider());

// ---------------- P0 ----------------
children.push(H1('المقدمة القابلة للعب — P0'));
children.push(Meta('المكان والزمان', 'شاشة سوداء، ثم غرفة النوم أمام المرآة — صباح آخر يوم دراسة'));
children.push(Meta('نقطة حفظ', 'لا يوجد؛ الحفظ الأول بعد C4'));

children.push(Stage('شاشة سوداء تمامًا. صوت إشعار هاتف واحد يكسر الصمت.'));
children.push(Narr('بعد خمس سنوات، ستتذكر بعض الأيام أكثر من غيرها.'));

children.push(Stage('واجهة اختيار: حقل اسم اللاعب، ثم اختيار الجنس (يغيّر الفن والتسجيل الصوتي فقط، دون أثر على الأحداث أو النتائج)'));
children.push(Narr('قبل أن تبدأ، اختر اسمك، واختر كيف تريد أن تُخاطَب في القصة.'));

children.push(Stage('انتقال من الأسود إلى غرفة نوم بسيطة، ضوء الصباح يدخل من الستارة، مرآة صغيرة على الحائط'));
children.push(Narr('تقف أمام المرآة في صباح اليوم الأخير من الدراسة. الحقيبة جاهزة عند الباب، والهاتف على السرير يعرض إشعارًا لم تفتحه بعد.'));

children.push(Line('السرد', 'لا تعرف بعد أي باب ستفتح. تعرف فقط أن اليوم يحمل أول قرار لن يستطيع أحد اتخاذه بدلًا منك.'));

children.push(Stage('انتقال بصري: من غرفة النوم إلى ممر المدرسة المزدحم بالطلاب في اليوم الأخير'));
children.push(transition('نهاية المقدمة. بداية الفصل الأول — C1: آخر جرس.'));
children.push(Divider());

// ---------------- C1 ----------------
children.push(H1('الفصل 1: آخر جرس — C1'));
children.push(Meta('المكان والزمان', 'الرياض — الأسبوع الأخير من المرحلة الثانوية، فصل دراسي ثم طريق العودة'));
children.push(Meta('الشخصيات', 'الأستاذ فهد'));

children.push(Stage('صف دراسي، الطلاب يجمعون كتبهم. الأستاذ فهد يقف عند اللوح ولا يغادر مكانه'));
children.push(Narr('ينتهي الدرس الأخير، لكن الأستاذ فهد يطلب من الطلاب البقاء لدقائق. يلتفت إلى اللوح ويكتب بخط واضح:'));
children.push(Stage('يكتب على اللوح: «أين تريد أن تكون بعد خمس سنوات؟» ثم يمسح كلمة «أين» بيده ويكتب مكانها «كيف»'));
children.push(Line('الأستاذ فهد', 'الخطة الجيدة لا تتنبأ بالمستقبل؛ تكشف فقط كيف ستتصرف عندما يتغير.'));

children.push(Stage('انتقال إلى طريق العودة سيرًا، شاشة الهاتف تضيء بإشعارين متتاليين'));
children.push(Narr('في طريق العودة يصل إلى هاتفك إشعاران: جدول مراجعة مكثف للاختبار النهائي، وعرض عمل موسمي مدفوع في مركز معارض. الوقت لا يسمح بكليهما بالصورة الكاملة.'));
children.push(Narr('هذه أول مرة تشعر فيها أن كل نعم تقولها لشيء تعني لا لشيء آخر.'));

children.push(decisionCard(
  'الأسبوعان القادمان',
  'أمامك أسبوعان قبل الاختبار النهائي. كيف تقضيهما؟',
  [
    {
      code: 'C1-A',
      button: 'تكوين مجموعة مراجعة والالتزام بجدول أسبوعي',
      response: 'تكتب اسمك في جدول المجموعة، وتشعر أن للمراجعة الآن شكلًا واضحًا.',
      tags: '[معرفة ▲ انضباط ▲]',
      dev: 'معرفة +8، انضباط +6، اتزان +2، ضغط +4. يفتح ذكرى «بدأت بالخطة» وحوارًا إضافيًا مع فهد لاحقًا.',
    },
    {
      code: 'C1-B',
      button: 'قبول العمل الموسمي وتعلّم التعامل مع الجمهور',
      response: 'توافق على الوردية، وتحفظ رقم المشرف في هاتفك.',
      tags: '[مال ▲ جرأة ▲]',
      dev: 'مال +8، جرأة +6، انضباط +3، ضغط +5. يفتح علم «أول_عمل» ويمنح خبرة عملية تُستخدم في فرص لاحقة.',
    },
  ],
));

children.push(transition('كلا الخيارين يقودان إلى الفصل الثاني (C2)، مع افتتاحية مختلفة حسب العلم المسجَّل.'));
children.push(Divider());

// ---------------- C2 ----------------
children.push(H1('الفصل 2: ثلاثة أبواب — C2'));
children.push(Meta('المكان والزمان', 'المنزل — مساء إعلان نتائج الثانوية'));

children.push(H3('افتتاحية مشروطة بعلم الفصل الأول'));
children.push(Stage('إذا لم يُسجَّل علم «أول_عمل»'));
children.push(Narr('أسابيع المراجعة انتهت أخيرًا، وها أنت تنتظر النتيجة وأنت أهدأ مما توقعت.'));
children.push(Stage('إذا سُجِّل علم «أول_عمل»'));
children.push(Narr('بين ورديات العمل الموسمي وأيام الانتظار، وصلت اللحظة التي كنت تؤجل التفكير فيها.'));

children.push(Narr('تظهر النتيجة على الشاشة، وبعدها تبدأ الرسائل والآراء. أمامك قبول جامعي في تخصص واسع، ومقعد في دبلوم تقني أقصر وأكثر تطبيقًا، وفرصة متدرب بدوام كامل لدى شركة محلية. لا يوجد باب يضمن النجاح، لكن لكل باب تكلفة لا تظهر في خطاب القبول.'));

children.push(Stage('تفتح دفتر الأستاذ فهد القديم، الذي أهداك إياه قبل التخرج'));
children.push(Line('دفتر الأستاذ فهد', 'لا تختَر الصورة التي تعجب الناس؛ اختر الثمن الذي تستطيع الالتزام به.'));
children.push(Narr('تتوقف أمام الخيارات الثلاثة، وتقرر أن يكون هذا القرار لك لا للضجيج المحيط بك.'));

children.push(decisionCard(
  'أي باب تفتح؟',
  'ثلاثة أبواب أمامك، ولا وقت لتجربتها كلها.',
  [
    {
      code: 'C2-A',
      button: 'الالتحاق بالجامعة وبناء أساس معرفي طويل',
      response: 'تحفظ رسالة القبول الجامعي، وتبدأ تتخيل أول محاضرة.',
      tags: '[معرفة ▲ مال ▼]',
      dev: 'معرفة +10، مال -5، ضغط +4. المسار = جامعة؛ تتغير مواقع ومشاهد C4 وC5.',
    },
    {
      code: 'C2-B',
      button: 'اختيار الدبلوم التقني والدخول المبكر إلى التطبيق',
      response: 'توقّع على مقعد الدبلوم، وتشعر أن الطريق أقصر وأوضح.',
      tags: '[انضباط ▲ معرفة ▲]',
      dev: 'معرفة +7، انضباط +8، جرأة +2، مال -2. المسار = دبلوم؛ يفتح لغة مهنية ومشاهد تدريب عملي.',
    },
    {
      code: 'C2-C',
      button: 'قبول الوظيفة والبدء بخبرة وراتب من الآن',
      response: 'تقبل عرض التدريب بدوام كامل، ويصلك أول موعد دوام.',
      tags: '[مال ▲ جرأة ▲]',
      dev: 'مال +10، جرأة +7، انضباط +4، معرفة -2، ضغط +6. المسار = عمل؛ يفتح خبرة سوق مبكرة وخيار دراسة مسائية لاحقًا.',
    },
  ],
));

children.push(transition('تلتقي المسارات الثلاثة في الفصل الثالث (C3)، مع بقاء متغيّر المسار (studyPath) فعالًا حتى النهاية.'));
children.push(Divider());

// ---------------- C3 ----------------
children.push(H1('الفصل 3: أول مبلغ يخصك — C3'));
children.push(Meta('المكان والزمان', 'بعد ثلاثة أشهر — تطبيق البنك ومتجر إلكتروني مفتوح على الهاتف'));

children.push(Narr('يتجمع لديك أول مبلغ تشعر أنه نتيجة قراراتك، سواء أتى من راتب تدريبك، أو مكافأة عملك، أو دخل من عمل جزئي بجانب دراستك. ثلاثة آلاف ريال ليست ثروة، لكنها تكفي لتكشف طريقة تفكيرك.'));
children.push(Narr('جهاز أفضل ودورة متخصصة قد يسرعان تعلمك. صندوق طوارئ يمنحك وقتًا عند التعثر. وتجربة متجر صغير قد تتحول إلى فرصة أو درس مكلف.'));
children.push(Narr('لا يظهر زر يحمل اسم «الخيار الصحيح». يظهر فقط رصيدك، وموعد الشهر القادم، وفكرة أنك ستتذكر هذه اللحظة عندما تصبح الأرقام أكبر.'));

children.push(decisionCard(
  'أول ثلاثة آلاف ريال',
  'الرصيد أمامك. ماذا تفعل به؟',
  [
    {
      code: 'C3-A',
      button: 'شراء أداة مناسبة ودورة مرتبطة بالمسار',
      response: 'يصلك الجهاز الجديد، وتفتح أول درس في الدورة.',
      tags: '[معرفة ▲ مال ▼]',
      dev: 'مال -8، معرفة +8، انضباط +3. علم «استثمار_مهاري»؛ يقلّل شرط المعرفة لبعض الفرص لاحقًا.',
    },
    {
      code: 'C3-B',
      button: 'تأسيس صندوق طوارئ وعدم لمسه',
      response: 'تفتح حسابًا منفصلًا وتودع فيه أول مبلغ، دون أن تلمسه.',
      tags: '[اتزان ▲ انضباط ▲]',
      dev: 'انضباط +6، اتزان +7. علم «صندوق_طوارئ»؛ يخفف أثر الأزمة المالية في C10.',
    },
    {
      code: 'C3-C',
      button: 'إطلاق تجربة بيع إلكتروني صغيرة بميزانية محدودة',
      response: 'تنشر أول منتج في متجرك الصغير، وتنتظر أول طلب.',
      tags: '[جرأة ▲ مال ▼]',
      dev: 'مال -7، جرأة +8، انضباط +2، ضغط +4. علم «مشروع_جانبي»؛ يغيّر قوة عرض الشراكة في C9.',
    },
  ],
));

children.push(transition('جميع الخيارات تقود إلى الفصل الرابع (C4)؛ تظهر نتيجة الاختيار كعنصر بصري في خلفية غرفة الشخصية وواجهة الرصيد.'));
children.push(Divider());

// ---------------- C4 ----------------
children.push(H1('الفصل 4: العلامة الحمراء — C4'));
children.push(Meta('المكان والزمان', 'نهاية العام الأول — قاعة دراسة أو ورشة تدريب أو مكتب، حسب المسار'));
children.push(Meta('نقطة حفظ', 'حفظ تلقائي عند نهاية هذا الفصل، ثم شاشة تطور بصري V1'));

children.push(H3('افتتاحية مشروطة بالمسار (studyPath)'));
children.push(Stage('إذا كان المسار = جامعة'));
children.push(Narr('تصل نتيجة مادة دراسية أقل مما توقعت — تعثر لم تكن تحسب له حسابًا.'));
children.push(Stage('إذا كان المسار = دبلوم'));
children.push(Narr('يصلك تقييم عملي من المدرب أقل من المستوى الذي اعتدت عليه.'));
children.push(Stage('إذا كان المسار = عمل'));
children.push(Narr('تصلك شكوى من عميل بسبب تفصيل صغير أغفلته في تسليمك الأخير.'));

children.push(Narr('العبارة مختلفة حسب مسارك، لكن الشعور واحد. تراجع ما حدث وتكتشف أن المشكلة ليست في القدرة وحدها؛ كانت هناك إشارات تجاهلتها لأنك أردت أن تبدو مسيطرًا.'));

children.push(Stage('إشعار رسالة قصيرة يصل على الهاتف من الأستاذ فهد'));
children.push(Line('الأستاذ فهد', 'التعثر معلومة. السؤال: هل ستقرأها أم ستخفيها؟'));
children.push(Narr('يمكنك طلب مساعدة تكشف ضعفك، أو حل المشكلة وحدك مهما طال الوقت، أو تغيير المسار قبل أن يستهلكك.'));

children.push(decisionCard(
  'العلامة الحمراء',
  'وصلتك أول نتيجة لا تريد رؤيتها. كيف تتعامل معها؟',
  [
    {
      code: 'C4-A',
      button: 'طلب مراجعة صريحة من فهد أو مشرف موثوق',
      response: 'ترسل رسالة لفهد، وتشعر بثقل يخف قليلًا بعد إرسالها.',
      tags: '[اتزان ▲ سمعة ▲]',
      dev: 'معرفة +6، اتزان +6، سمعة +4، ضغط -4. علم «مرشد»؛ يفتح استشارة إضافية عند أزمة C10.',
    },
    {
      code: 'C4-B',
      button: 'إعادة العمل بمفردك حتى ينجح دون إخبار أحد',
      response: 'تغلق الباب وتبدأ من جديد، وحدك، حتى ينجح الأمر.',
      tags: '[انضباط ▲ ضغط ▲]',
      dev: 'انضباط +7، جرأة +3، ضغط +8. علم «حل_منفرد»؛ يمنح إنجازًا سريعًا لكنه يزيد الضغط المتراكم.',
    },
    {
      code: 'C4-C',
      button: 'إعادة تقييم المسار والانتقال إلى خيار أقرب لقدراتك',
      response: 'تعيد ترتيب أولوياتك وتبدّل مسارك القريب دون إعلان صاخب.',
      tags: '[جرأة ▲ سمعة ▼]',
      dev: 'جرأة +9، اتزان +3، مال -3، سمعة -2. علم «تحول»؛ يغيّر وصف السيرة في C5 دون عقوبة دائمة.',
    },
  ],
));

children.push(transition('نقطة التقاء أولى، ثم شاشة تطور بصري (V1) تعرض شكل الشخصية الجديد، قبل بداية الفصل الخامس (C5) في الدفعة التالية.'));
children.push(Divider());

// ---------------- CLOSING ----------------
children.push(H1('حالة الدفعة والخطوة التالية'));
children.push(P([rtlRun({ text: 'هذا المستند يغطي: ', size: 22 }), rtlRun({ text: 'المقدمة P0، والفصول C1–C4 كاملة، مع كل الخيارات والاستجابات الفورية والانتقالات ونقاط الحفظ.', size: 22, bold: true })]));
children.push(P('المتبقي حسب خطة الإنتاج: كتابة الحوار الكامل للفصول C5–C12 ومحرك النهايات E1–E4 بالأسلوب نفسه، عند الرغبة في المتابعة.'));

const doc = new Document({
  styles: {
    default: { document: { run: { font: FONT, size: 22 }, paragraph: { bidirectional: true, alignment: AlignmentType.RIGHT } } },
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 },
        },
      },
      children,
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  fs.writeFileSync('/home/user/workspace/masarak_dialogue_p0_c1_c4.docx', buf);
  console.log('done');
});
