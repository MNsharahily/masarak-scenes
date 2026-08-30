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

const ActHeader = (text) => new Paragraph({
  alignment: AlignmentType.RIGHT,
  bidirectional: true,
  spacing: { before: 560, after: 260 },
  shading: { fill: ACCENT_DARK, type: ShadingType.CLEAR },
  children: [rtlRun({ text, size: 30, bold: true, color: 'FFFFFF' })],
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

function decisionCard(promptTitle, promptText, options) {
  const cellMargins = { top: 100, bottom: 100, left: 150, right: 150 };
  const borderAll = { style: BorderStyle.SINGLE, size: 2, color: BORDER };
  const borders = { top: borderAll, bottom: borderAll, left: borderAll, right: borderAll, insideHorizontal: borderAll, insideVertical: borderAll };
  const tableWidth = 9360;

  const rows = [];
  rows.push(new TableRow({
    tableHeader: true,
    children: [new TableCell({
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

  return new Table({ width: { size: tableWidth, type: WidthType.DXA }, columnWidths: [tableWidth], borders, rows });
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
  new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, children: [rtlRun({ text: 'مَسَارُك', size: 64, bold: true, color: ACCENT_DARK })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { before: 160 }, children: [rtlRun({ text: 'الحوار الكامل للشاشة — الجزء الثاني', size: 34, bold: true, color: ACCENT })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { before: 80 }, children: [rtlRun({ text: 'الفصول من الخامس إلى الثاني عشر (C5–C12)', size: 24, color: TEXT })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { before: 40 }, children: [rtlRun({ text: 'يكمل هذا الجزء «الحوار الكامل للشاشة — المقدمة والفصول C1–C4»، بالأسلوب والقواعد نفسها', size: 20, italics: true, color: MUTED })] }),
  new Paragraph({ spacing: { before: 600 }, children: [] }),
);

// ---------------- SECTION 0 ----------------
children.push(H1('0. تذكير سريع بقواعد هذه الدفعة'));
[
  'صيغة مخاطب محايدة، مضارع مستمر، وتجنّب لواحق التذكير والتأنيث حيث لا يلزم.',
  'نص كل زر بين 4 و10 كلمات، يبدأ بفعل، ولا يكشف الأثر الرقمي.',
  'بعد كل اختيار: جملة واحدة كاستجابة فورية، بلا حكم أخلاقي وبلا تشخيص.',
  'رموز السمات (للفريق فقط): K معرفة، D انضباط، C جرأة، B اتزان، R سمعة، M مال، P ضغط.',
  'مسارات كل فصل تُكتب كطبقة نص شرطي قصيرة فوق بنية مشتركة، لا كقصص منفصلة.',
].forEach(t => children.push(new Paragraph({
  alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { after: 80 },
  children: [rtlRun({ text: '•  ', size: 22, color: ACCENT }), rtlRun({ text: t, size: 22, color: TEXT })],
})));
children.push(Divider());

// ================= ACT 2: الاستقلال =================
children.push(ActHeader('القسم السردي الثاني: الاستقلال — الفصول C5 إلى C8'));

// ---------------- C5 ----------------
children.push(H1('الفصل 5: فرصة لا تأتي مرتين؟ — C5'));
children.push(Meta('المكان والزمان', 'بعد عامين — بوابة فرص، مركز تدريب، أو اجتماع تقييم، حسب المسار'));

children.push(H3('افتتاحية مشروطة بالمسار (studyPath)'));
children.push(Stage('إذا كان المسار = جامعة'));
children.push(Narr('يعرض عليك مشرفك فرصة تدريب نوعي، أو التركيز الكامل على مسار أكاديمي متقدم.'));
children.push(Stage('إذا كان المسار = دبلوم'));
children.push(Narr('يصلك عرض شهادة تخصصية إضافية، أو عقد عمل مباشر من جهة تدرّبت لديها.'));
children.push(Stage('إذا كان المسار = عمل'));
children.push(Narr('تُعرض عليك ترقية داخلية، أو فرصة دراسة مسائية توازي عملك الحالي.'));

children.push(Narr('تمر سنتان. لم تعد الشخصية تبدو كما بدأت: بطاقة جامعية أو زي تدريب أو شارة عمل، وحركة أكثر ثقة. تصل فرصة ترفع سقف المرحلة، لكنها تحمل العبارة المعتادة: «نحتاج ردك سريعًا.» السرعة تجعل العرض يبدو نادرًا، مع أنه ليس واضحًا إن كان مناسبًا فعلًا.'));
children.push(Narr('الخيار الأول يعمق تخصصك لكنه يؤخر الدخل. الثاني يضعك في عمل أكثر ومسؤولية أكبر. الثالث يحاول الجمع بينهما، وهو ممكن، لكن ليس بلا ثمن على الوقت والتركيز.'));

children.push(decisionCard(
  'فرصة لا تنتظر',
  'العرض أمامك، ولا وقت طويلاً للتفكير. ماذا تختار؟',
  [
    { code: 'C5-A', button: 'تعميق التخصص وتأجيل العائد القريب', response: 'ترفض العرض السريع وتختار برنامجًا أعمق يبطئ دخلك قليلًا.', tags: '[معرفة ▲ مال ▼]', dev: 'معرفة +10، انضباط +5، مال -4، ضغط +4. علم «متخصص»؛ يقوي نهاية الخبير.' },
    { code: 'C5-B', button: 'قبول فرصة الدخل والمسؤولية المبكرة', response: 'توقّع فورًا، وتشعر أن السقف ارتفع بسرعة لم تتوقعها.', tags: '[مال ▲ جرأة ▲]', dev: 'مال +11، جرأة +6، سمعة +2، ضغط +6. علم «مسار_سريع»؛ يسرّع العروض التجارية لاحقًا.' },
    { code: 'C5-C', button: 'الجمع بين الدراسة والعمل بخطة زمنية صارمة', response: 'تضع جدولًا صارمًا يجمع بين الاثنين، وتعرف أنه سيكلفك راحتك.', tags: '[انضباط ▲ ضغط ▲]', dev: 'معرفة +5، انضباط +5، مال +5، ضغط +10. علم «مسار_هجين»؛ مكسب مزدوج مع خطر إرهاق واضح في C8.' },
  ],
));
children.push(transition('جميع المسارات تتقاطع في الفصل السادس (C6) بعد قفزة زمنية قصيرة.'));
children.push(Divider());

// ---------------- C6 ----------------
children.push(H1('الفصل 6: خمسة عشر ألفًا — C6'));
children.push(Meta('المكان والزمان', 'ليلة هادئة — لوحة الميزانية الشخصية'));
children.push(Meta('الشخصيات', 'سارة، راكان'));

children.push(Narr('بالتدرج والالتزام يصل ما تملكه إلى مبلغ يسمح بقرار حقيقي. على الشاشة ثلاثة ملفات: اعتماد مهني يرفع قيمتك، احتياطي يغطي شهورًا صعبة، وفكرة متجر رقمي طورها صديقك راكان.'));
children.push(Line('راكان', 'إذا انتظرنا حتى نضمن كل شيء، لن نبدأ.'));
children.push(Line('سارة', 'وإذا بدأنا بلا حساب، سندفع ثمن الحماس.'));
children.push(Narr('المشهد لا يسخر من المخاطرة ولا يمجد الادخار. المطلوب أن تختار ما يناسب المسار الذي بنيته، وأن تتحمل ما يأتي معه.'));

children.push(decisionCard(
  'خمسة عشر ألف ريال',
  'أمامك ثلاثة استخدامات لهذا المبلغ. أيها تختار؟',
  [
    { code: 'C6-A', button: 'تمويل اعتماد مهني قوي مرتبط بتخصصك', response: 'تدفع رسوم الاعتماد بالكامل، وتبدأ التحضير في الليلة نفسها، وتشعر أن اسمك المهني بدأ يكتسب وزنًا أكبر بين زملائك.', tags: '[معرفة ▲ سمعة ▲ مال ▼]', dev: 'مال -10، معرفة +11، انضباط +4، سمعة +3 (الاعتماد المهني يمنح مصداقية إضافية بين الزملاء). علم «معتمد»؛ يفتح حوار خبير وفرصة قيادية فنية في C12.' },
    { code: 'C6-B', button: 'تعزيز الاحتياطي ووضع قواعد واضحة للصرف', response: 'تنقل جزءًا أكبر من المبلغ إلى الاحتياطي، وتكتب لنفسك قاعدة صرف واضحة.', tags: '[اتزان ▲ انضباط ▲]', dev: 'انضباط +6، اتزان +9، ضغط -3. علم «احتياطي_قوي»؛ حماية إضافية في أزمة C10.' },
    { code: 'C6-C', button: 'تمويل نسخة محدودة من متجر راكان واختبار السوق', response: 'توافق على تمويل نسخة تجريبية من فكرة راكان، وتصبح شريكًا صغيرًا فيها.', tags: '[جرأة ▲ مال ▼]', dev: 'مال -9، جرأة +10، انضباط +3، ضغط +6. علم «بذرة_مشروع»؛ يقوي عرض C9 ويضيف عميلًا مبكرًا.' },
  ],
));
children.push(transition('الانتقال إلى الفصل السابع (C7)؛ لا تظهر النتيجة المالية الكبرى إلا لاحقًا حتى يبقى القرار ذا توتر.'));
children.push(Divider());

// ---------------- C7 ----------------
children.push(H1('الفصل 7: رقم واحد زائد — C7'));
children.push(Meta('المكان والزمان', 'المكتب — قبل إرسال تقرير أداء إلى عميل مهم'));
children.push(Meta('الشخصيات', 'مها'));

children.push(Narr('قبل دقائق من إرسال التقرير تلاحظ أن لوحة البيانات كررت مجموعة من السجلات. الرقم الظاهر تسعون، والرقم الصحيح ستون. التقرير القوي قد يفتح عقدًا أو ترقية، والتصحيح الآن قد يؤخر الاجتماع ويكشف أن المراجعة تمت في اللحظة الأخيرة.'));
children.push(Line('مها', 'أحتاج النسخة النهائية خلال عشر دقائق.'));
children.push(Narr('يفتح المؤشر فوق زر الإرسال. القضية ليست فقط صدقًا أو كذبًا؛ إنها طريقة تحمل الخطأ، وتوقيت المواجهة، ومن سيدفع ثمن قرارك.'));

children.push(decisionCard(
  'رقم واحد زائد',
  'الرقم في التقرير خاطئ، والوقت ينفد. ماذا تفعل قبل الإرسال؟',
  [
    { code: 'C7-A', button: 'إيقاف الإرسال، تصحيح الرقم، وإبلاغ مها على انفراد', response: 'توقف الإرسال في اللحظة الأخيرة، وتطرق باب مها بنفسك.', tags: '[سمعة ▲ اتزان ▲]', dev: 'سمعة +10، انضباط +4، اتزان +4، مال -2. علم «شفافية»؛ تستشهد مها بهذا الموقف في C11.' },
    { code: 'C7-B', button: 'إرسال التقرير ثم محاولة تعديل الرقم لاحقًا', response: 'تضغط زر الإرسال، وتخطط لتصحيح الرقم بهدوء قبل أن يسأل أحد.', tags: '[مال ▲ سمعة ▼]', dev: 'مال +5، جرأة +3، سمعة -8، ضغط +9. علم «خطأ_مخفي»؛ قد يعود في التفاوض أو في نهاية اللعبة.' },
    { code: 'C7-C', button: 'إعلان الخطأ أمام الجميع وتحميل زميل مسؤولية البيانات', response: 'تعلن الخطأ في الاجتماع، وتوجّه أصابع الاتهام نحو زميلك المسؤول عن البيانات.', tags: '[جرأة ▲ سمعة ▼]', dev: 'جرأة +6، سمعة -3، انضباط -2، ضغط +6. علم «صدام_علني»؛ يغيّر علاقة الفريق في C11.' },
  ],
));
children.push(transition('جميع الخيارات تقود إلى الفصل الثامن (C8)، مع اختلاف نبرة مها وثقة الفريق في المشاهد التالية.'));
children.push(Divider());

// ---------------- C8 ----------------
children.push(H1('الفصل 8: السنة التي بلا فراغ — C8'));
children.push(Meta('المكان والزمان', 'نهاية عام مزدحم — تقويم ممتلئ ومكالمة من مها'));
children.push(Meta('نقطة حفظ', 'حفظ تلقائي عند نهاية هذا الفصل، ثم شاشة تطور بصري V2'));

children.push(Narr('يبدو العام ناجحًا من الخارج: مهام أكثر، دخل أفضل، ورسائل تبدأ بعبارة «نحتاجك». في الداخل، لم يعد التقويم يحتوي على مساحة بيضاء. تعرض مها مسؤولية أكبر مقابل ساعات أطول. في الوقت نفسه يفتح التسجيل في برنامج تخصصي انتظرته، وتلاحظ أن تركيزك أصبح أقصر من المعتاد.'));
children.push(Narr('الطموح يطلب منك ألا تفوت الفرصة. الخبرة تقول إن الجسد والوقت موارد أيضًا. يمكنك الدفع بقوة، أو الاستثمار في المعرفة، أو إعادة ترتيب حياتك قبل الخطوة التالية.'));

children.push(decisionCard(
  'عام بلا فراغ',
  'التقويم ممتلئ، والخيارات الثلاثة كلها مغرية بطريقتها. أي طريق تسلك؟',
  [
    { code: 'C8-A', button: 'قبول الترقية والعمل الإضافي لعام كامل', response: 'توقّع على الترقية، وتحذف كلمة «إجازة» من تقويمك لعام كامل.', tags: '[مال ▲ ضغط ▲]', dev: 'مال +13، جرأة +8، انضباط +4، ضغط +13، اتزان -6. علم «ترقية»؛ يرفع قوة التفاوض لاحقًا لكنه يجعل الضغط ظاهرًا.' },
    { code: 'C8-B', button: 'رفض التوسع المؤقت والالتحاق بالبرنامج التخصصي', response: 'تعتذر عن التوسع مؤقتًا، وتسجل في البرنامج الذي انتظرته طويلًا.', tags: '[معرفة ▲ مال ▼]', dev: 'معرفة +13، انضباط +7، مال -7، ضغط +5. علم «دراسة_متقدمة»؛ يقوي نهاية الخبير وفرص C12.' },
    { code: 'C8-C', button: 'تخفيض الحمل وبناء روتين قابل للاستمرار', response: 'تعيد ترتيب أسبوعك بالكامل، وتترك مساحة بيضاء في التقويم أول مرة منذ طويل.', tags: '[اتزان ▲ ضغط ▼]', dev: 'اتزان +14، انضباط +4، ضغط -12، مال -3. علم «إعادة_ضبط»؛ يمنح خيارات تفاوض هادئة في أزمة C10.' },
  ],
));
children.push(transition('نقطة التقاء ثانية وشاشة تطور بصري (V2)، ثم بداية الفصل التاسع (C9) في القسم السردي الثالث.'));
children.push(Divider());

// ================= ACT 3: المسؤولية =================
children.push(ActHeader('القسم السردي الثالث: المسؤولية — الفصول C9 إلى C12'));

// ---------------- C9 ----------------
children.push(H1('الفصل 9: عرض على طاولة صغيرة — C9'));
children.push(Meta('المكان والزمان', 'مقهى عمل مشترك — اجتماع مع سارة وراكان'));
children.push(Meta('الشخصيات', 'سارة، راكان'));

children.push(H3('افتتاحية مشروطة بالأعلام السابقة'));
children.push(Stage('إذا وُجد علم «مشروع_جانبي» أو «بذرة_مشروع»'));
children.push(Narr('يفتح راكان الحاسوب ويعرض قائمة عملاء يعرفون تجربتك السابقة بالفعل، ويسألونك متى تبدؤون رسميًا.'));
children.push(Stage('إذا كان المسار = عمل ولم توجد الأعلام أعلاه (مسار مستقر)'));
children.push(Narr('قبل الاجتماع بساعة، يصلك على الهاتف عرض ترقية بديل من عملك الحالي، وكأن التوقيت مقصود.'));

children.push(Narr('تضع سارة نموذجًا أوليًا لمنصة تربط الطلاب بفرص تدريب قصيرة وموثقة. لدى راكان قائمة عملاء محتملين، ولديك خبرة من المسار الذي سلكته.'));
children.push(Line('راكان', 'السوق لن ينتظر.'));
children.push(Narr('تقترح سارة اختبارًا محدودًا قبل التوسع. في الجهة الأخرى، وظيفتك الحالية أصبحت أكثر استقرارًا وتمنحك مسارًا واضحًا.'));
children.push(Narr('الشراكة السريعة قد تضاعف الدخل أو تضاعف الأخطاء. التجربة المحدودة أبطأ لكنها تمنح بيانات حقيقية. والرفض ليس خوفًا بالضرورة؛ قد يكون اختيارًا واعيًا للاستقرار أو التخصص.'));

children.push(decisionCard(
  'عرض على طاولة صغيرة',
  'سارة وراكان ينتظران ردك. ماذا تقرر؟',
  [
    { code: 'C9-A', button: 'دخول الشراكة والتوسع السريع من البداية', response: 'توافق على الشراكة الكاملة، وتبدأ التوسع قبل أن تجهز الأرض تمامًا.', tags: '[جرأة ▲ ضغط ▲]', dev: 'جرأة +11، مال +7، سمعة +3، ضغط +12. علم «شراكة_سريعة»؛ أزمة C10 تصبح تأخر مستحقات كبير (C10-F).' },
    { code: 'C9-B', button: 'تشغيل تجربة لمدة تسعين يومًا بمؤشرات نجاح واضحة', response: 'تقترح تجربة محدودة بمؤشرات واضحة، ويوافق الطرفان على تسعين يومًا فقط.', tags: '[انضباط ▲ سمعة ▲]', dev: 'معرفة +5، انضباط +9، جرأة +5، مال -5، سمعة +5، ضغط +5. علم «تجربة_محدودة»؛ أزمة C10 تصبح خطأ مورد (C10-P)، ويقوي نهاية الأثر المستدام.' },
    { code: 'C9-C', button: 'رفض العرض والاستمرار في المسار المهني المستقر', response: 'تشكر سارة وراكان، وتختار البقاء في مسارك الحالي الأكثر وضوحًا.', tags: '[اتزان ▲ مال ▲]', dev: 'اتزان +10، مال +7، انضباط +3، جرأة -2، ضغط -5. علم «استقرار»؛ أزمة C10 تصبح إعادة هيكلة وظيفية (C10-S).' },
  ],
));
children.push(transition('C9-A تقود إلى C10-F، وC9-B تقود إلى C10-P، وC9-C تقود إلى C10-S؛ ثم تتقاطع الفروع الثلاثة مجددًا بعد الأزمة.'));
children.push(Divider());

// ---------------- C10 ----------------
children.push(H1('الفصل 10: التحويل المتأخر — C10-F / C10-P / C10-S'));
children.push(Meta('المكان والزمان', 'صباح الأحد — إشعار مالي لم يصل'));

children.push(H3('افتتاحية الأزمة حسب مسار الفصل التاسع'));
children.push(Stage('C10-F — مسار الشراكة السريعة (من C9-A)'));
children.push(Narr('الساعة العاشرة واثنتا عشرة دقيقة. تفتح الحساب مرة أخرى، ولا يظهر التحويل المتوقع: العميل الأكبر يؤجل السداد تسعين يومًا كاملة.'));
children.push(Stage('C10-P — مسار التجربة المحدودة (من C9-B)'));
children.push(Narr('الساعة العاشرة واثنتا عشرة دقيقة. رسالة من مورّدك تصل بدل التحويل: خطأ في الشحنة يتسبب في إعادة جزء كامل من العمل على حسابك.'));
children.push(Stage('C10-S — مسار الاستقرار (من C9-C)'));
children.push(Narr('الساعة العاشرة واثنتا عشرة دقيقة. بريد رسمي من الإدارة بدل إشعار الراتب المعتاد: الشركة تعلن إعادة هيكلة وتجميد المكافآت هذا الربع.'));

children.push(Narr('الأزمة ليست نهاية القصة؛ إنها اللحظة التي تكشف قيمة احتياطيك، وسمعتك، وطريقة تعاملك مع الخوف. يمكنك زيادة الرهان بالتمويل، أو التفاوض وتقليص النطاق، أو الخروج مبكرًا وحماية ما بقي.'));
children.push(Stage('نص شرطي معدِّل: إذا كان لديك علم «صندوق_طوارئ» أو «احتياطي_قوي»، تخف الخسارة والضغط الناتجان عن أي خيار أدناه. إذا كان لديك علم «مرشد»، تُضاف مكالمة قصيرة مع الأستاذ فهد قبل القرار.'));

children.push(decisionCard(
  'صباح بلا تحويل',
  'الأزمة وصلت. كيف تتعامل معها؟',
  [
    { code: 'C10-A', button: 'الحصول على تمويل قصير وزيادة الرهان لتجاوز الأزمة', response: 'توقّع على تمويل عاجل يغطي الفجوة، وتؤجل القلق إلى الشهر القادم.', tags: '[مال ▲ ضغط ▲]', dev: 'جرأة +9، مال +12، انضباط -3، ضغط +15. علم «دين»؛ يزيد العائد المحتمل وخطورة نهاية الضغط.' },
    { code: 'C10-B', button: 'التفاوض، تقليص النطاق، وجدولة الالتزامات بوضوح', response: 'تتصل بالطرف الآخر مباشرة، وتخرجان معًا بجدول أصغر لكنه واقعي.', tags: '[اتزان ▲ سمعة ▲]', dev: 'انضباط +7، اتزان +9، سمعة +5، مال -5، ضغط -6. علم «إعادة_هيكلة»؛ يقوي نهاية الأثر المستدام.' },
    { code: 'C10-C', button: 'إنهاء المشروع أو الدور سريعًا وحماية المتبقي', response: 'تغلق الملف بسرعة، وتحمي ما تبقى بدل أن تحاول إنقاذ كل شيء.', tags: '[اتزان ▲ ضغط ▼]', dev: 'اتزان +6، جرأة +4، مال -3، سمعة -4، ضغط -10. علم «خروج»؛ يفتح لغة «البداية الثانية» دون اعتباره فشلًا.' },
  ],
));
children.push(transition('تلتقي الفروع الثلاثة في الفصل الحادي عشر (C11) مع الاحتفاظ بعلم نوع الأزمة وطريقة إدارتها.'));
children.push(Divider());

// ---------------- C11 ----------------
children.push(H1('الفصل 11: اسمك على الخطأ — C11'));
children.push(Meta('المكان والزمان', 'اجتماع عميل — عرض نهائي أُدرج فيه سعر خاطئ'));
children.push(Meta('الشخصيات', 'مها، نوف'));

children.push(H3('افتتاحية مشروطة بعلم الفصل السابع'));
children.push(Stage('إذا وُجد علم «شفافية» (من C7-A)'));
children.push(Narr('تلاحظ مها التوتر في الغرفة، وتمنحك مساحة إضافية قبل أن توجّه سؤالها.'));
children.push(Stage('إذا وُجد علم «خطأ_مخفي» أو «صدام_علني» (من C7-B أو C7-C)'));
children.push(Narr('يبدو العميل أكثر تشككًا من المعتاد، وتصبح كلفة أي قرار هنا أعلى مما كانت لتكون.'));

children.push(Narr('قبل العرض بدقائق تكتشف نوف، أحدث أعضاء الفريق، أن السعر في النسخة المرسلة أقل من التكلفة الفعلية. الخطأ بدأ منها، لكنه مر عبر موافقتك.'));
children.push(Line('مها', 'من اعتمد هذه النسخة؟'));
children.push(Narr('يصمت الجميع، وتنظر نوف إلى الطاولة.'));
children.push(Narr('إن تحملت المسؤولية ستدفع تكلفة مباشرة وقد تكسب فريقًا يثق بك. إن نسبت الخطأ إلى نوف قد تحمي الصفقة مؤقتًا. وإن أخفيت الفرق داخل بند آخر فقد لا يكتشفه أحد اليوم. هنا يتحدد نوع القائد الذي أصبحت عليه.'));

children.push(decisionCard(
  'اسمك على الخطأ',
  'مها تنتظر جوابًا أمام الجميع. من يتحمل هذا الخطأ؟',
  [
    { code: 'C11-A', button: 'تحمل المسؤولية، تصحيح العرض، وتدريب نوف بعد الاجتماع', response: 'تقول أمام الجميع: «الموافقة كانت مني»، ثم تجلس مع نوف بعد الاجتماع لتشرح لها الفرق.', tags: '[سمعة ▲ اتزان ▲]', dev: 'سمعة +12، اتزان +7، انضباط +4، مال -3، ضغط +2. علم «مسؤولية»؛ شرط قوي لنهاية الأثر المستدام.' },
    { code: 'C11-B', button: 'تحديد نوف بوصفها المسؤولة وحماية موقعك', response: 'تشير إلى نوف بهدوء، وتترك الاجتماع وموقعك سليم لكن الغرفة أكثر بردًا.', tags: '[مال ▲ سمعة ▼]', dev: 'مال +5، جرأة +4، سمعة -10، ضغط +7. علم «لوم»؛ يضعف ثقة الفريق ويغيّر مشهد النهاية.' },
    { code: 'C11-C', button: 'إخفاء الفرق في بند آخر وإكمال الصفقة', response: 'تعيد توزيع الفرق داخل بند آخر بهدوء، وتغلق الصفقة قبل أن يلاحظ أحد.', tags: '[مال ▲ سمعة ▼]', dev: 'مال +7، سمعة -15، ضغط +11، انضباط -3. علم «إخفاء_فريق»؛ يرفع احتمال النهاية الهشة أو البداية الثانية.' },
  ],
));
children.push(transition('الانتقال إلى الفصل الثاني عشر (C12)؛ يعود سلوك C7 كمرآة لمدى تغيّر الشخصية أو تكرارها للنمط نفسه.'));
children.push(Divider());

// ---------------- C12 ----------------
children.push(H1('الفصل 12: المصعد الزجاجي — C12'));
children.push(Meta('المكان والزمان', 'برج أعمال في الرياض — بعد اجتماع الفرصة الأكبر'));
children.push(Meta('نقطة حفظ', 'حفظ تلقائي عند نهاية هذا الفصل، ثم شاشة تطور بصري V3 ومحرك النهايات'));

children.push(Stage('باب المصعد الزجاجي ينغلق ببطء، منظر الرياض يتسع خلف الزجاج'));
children.push(Narr('ينغلق باب المصعد بعد اجتماع قد يغير السنوات القادمة. العرض الأول يمول توسعًا سريعًا إلى عدة مدن مقابل التزامات كبيرة. الثاني ينمو ببطء، يدرب الفريق، ويقبل أرباحًا أقل في البداية. والثالث يتيح لك التراجع من الإدارة اليومية والتخصص أو الدراسة بعمق.'));

children.push(Stage('انعكاس في زجاج المصعد يعرض ثلاث لقطات سريعة من الماضي'));
children.push(Narr('في انعكاس الزجاج ترى النسخ السابقة منك: طالبًا أمام آخر اختبار، وشخصًا يراقب أول مبلغ في الحساب، وقائدًا أمام خطأ فريقه. لا تختار بين النجاح والفشل. تختار شكل الثمن الذي ستدفعه، ونوع الأثر الذي تريد أن يبقى بعدك.'));

children.push(decisionCard(
  'المصعد الزجاجي',
  'ثلاثة أشكال للمستقبل أمامك. أيها تختار؟',
  [
    { code: 'C12-A', button: 'قبول التمويل والتوسع السريع في عدة مدن', response: 'توقّع على التمويل، وتفتح خرائط لثلاث مدن جديدة في الأسبوع نفسه.', tags: '[مال ▲ ضغط ▲]', dev: 'جرأة +13، مال +16، سمعة +3، ضغط +17، اتزان -7. علم «توسع_سريع»؛ يرجّح نهاية الرائد الجريء (E2).' },
    { code: 'C12-B', button: 'نمو مستدام، تدريب الفريق، وتوسع على مراحل', response: 'تختار خطة أبطأ، وتبدأ بتدريب من يخلفك قبل أي توسع جديد.', tags: '[سمعة ▲ اتزان ▲]', dev: 'انضباط +9، معرفة +7، سمعة +10، مال +7، اتزان +5، ضغط +3. علم «استدامة»؛ يرجّح نهاية صانع الأثر (E1).' },
    { code: 'C12-C', button: 'التخصص العميق وتسليم الإدارة لمن هو أنسب', response: 'تسلّم مقاليد الإدارة اليومية، وتفتح لنفسك بابًا للتخصص العميق الذي أجّلته طويلًا.', tags: '[معرفة ▲ اتزان ▲]', dev: 'معرفة +14، اتزان +11، انضباط +5، مال -5، ضغط -10. علم «تمكّن»؛ يرجّح نهاية الخبير الموثوق (E3).' },
  ],
));
children.push(transition('الانتقال إلى محرك النهايات E1–E4، ثم عرض تقرير النمو السردي.'));
children.push(Divider());

// ---------------- CLOSING ----------------
children.push(H1('حالة الدفعة والخطوة التالية'));
children.push(P([rtlRun({ text: 'هذا المستند يغطي: ', size: 22 }), rtlRun({ text: 'الفصول من الخامس إلى الثاني عشر (C5–C12) كاملة، مع الافتتاحيات المشروطة، وكل الخيارات والاستجابات الفورية والانتقالات ونقاط الحفظ V2 وV3.', size: 22, bold: true })]));
children.push(P('بذلك تكتمل القصة الأساسية من المقدمة حتى نهاية C12. المتبقي حسب خطة الإنتاج: كتابة محرك النهايات الأربع (E1–E4) وتقرير النمو السردي الختامي بالأسلوب نفسه، عند الرغبة في المتابعة.'));

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
  fs.writeFileSync('/home/user/workspace/masarak_dialogue_c5_c12.docx', buf);
  console.log('done');
});
