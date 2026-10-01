const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType,
  AlignmentType, BorderStyle, ShadingType, LevelFormat, HeadingLevel, Header, Footer,
  PageNumber, PageBreak, VerticalAlign
} = require('docx');

const FONT = 'Times New Roman';
const W = 9071; // Chiều rộng nội dung (A4, lề trái 3cm, phải 2cm)
const BODY = 26; // 13pt
const TBL = 22;  // 11pt trong bảng

// ---------- Helpers ----------
function runs(str, base = {}) {
  const parts = str.split(/(\*\*[^*]+\*\*)/g).filter(s => s !== '');
  return parts.map(s =>
    s.startsWith('**')
      ? new TextRun({ font: FONT, size: BODY, ...base, text: s.slice(2, -2), bold: true })
      : new TextRun({ font: FONT, size: BODY, ...base, text: s })
  );
}

function para(str, o = {}) {
  return new Paragraph({
    children: runs(str, o.run || {}),
    alignment: o.align ?? AlignmentType.JUSTIFIED,
    spacing: { after: o.after ?? 100, before: o.before ?? 0, line: 312 },
    indent: o.indent,
    keepNext: o.keepNext,
  });
}

const say = (name, text) => para(`**${name}:** ${text}`, { indent: { left: 284 } });

const bullet = (str, lvl = 0, size = BODY) =>
  new Paragraph({
    numbering: { reference: 'bul', level: lvl },
    children: runs(str, { size }),
    alignment: AlignmentType.JUSTIFIED,
    spacing: { after: 60, line: 300 },
  });

const spacer = () => new Paragraph({ children: [], spacing: { after: 120 } });

function h(text, level) {
  const map = { 1: HeadingLevel.HEADING_1, 2: HeadingLevel.HEADING_2, 3: HeadingLevel.HEADING_3 };
  return new Paragraph({ heading: map[level], children: [new TextRun({ text, font: FONT })], keepNext: true });
}

const bd = { style: BorderStyle.SINGLE, size: 4, color: '808080' };
const borders = { top: bd, bottom: bd, left: bd, right: bd };

function cell(content, w, o = {}) {
  const lines = Array.isArray(content) ? content : String(content).split('\n');
  return new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders,
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: 'auto' } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    verticalAlign: VerticalAlign.TOP,
    children: lines.map(l =>
      new Paragraph({
        children: runs(l, { size: o.size ?? TBL, bold: o.bold }),
        alignment: o.align ?? AlignmentType.LEFT,
        spacing: { after: 40, line: 264 },
      })
    ),
  });
}

function table(widths, head, rows, o = {}) {
  const total = widths.reduce((a, b) => a + b, 0);
  const trs = [];
  if (head) {
    trs.push(new TableRow({
      tableHeader: true, cantSplit: true,
      children: head.map((t, i) => cell(t, widths[i], { fill: 'D9E2F3', bold: true, align: AlignmentType.CENTER, size: o.size })),
    }));
  }
  rows.forEach(r => trs.push(new TableRow({
    cantSplit: true,
    children: r.map((t, i) => cell(t, widths[i], {
      size: o.size,
      fill: o.labelCol && i === 0 ? 'F2F2F2' : undefined,
      bold: o.labelCol && i === 0,
      align: (o.center || []).includes(i) ? AlignmentType.CENTER : AlignmentType.LEFT,
    })),
  })));
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths, rows: trs });
}

function info(rows) { return table([2400, W - 2400], null, rows, { labelCol: true }); }

function concl(items, title = 'Kết luận') {
  const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: [W],
    rows: [new TableRow({
      children: [new TableCell({
        width: { size: W, type: WidthType.DXA },
        borders: { top: none, bottom: none, right: none, left: { style: BorderStyle.SINGLE, size: 24, color: '2E7D32' } },
        shading: { fill: 'EAF4EA', type: ShadingType.CLEAR, color: 'auto' },
        margins: { top: 80, bottom: 60, left: 160, right: 120 },
        children: [
          new Paragraph({ children: [new TextRun({ text: title + ':', bold: true, font: FONT, size: BODY })], spacing: { after: 60 } }),
          ...items.map(s => bullet(s)),
        ],
      })],
    })],
  });
}

// ---------- Nội dung cuộc họp ----------
function meeting(base) {
  const sec = t => h(t, base);
  const sub = t => h(t, base + 1);
  const c = [];

  c.push(info([
    ['Học phần', 'KNM – Hoạt động nhóm tuần 3'],
    ['Giảng viên', 'TS. Lê Lý Thùy Trâm (Bộ môn Công nghệ sinh học, Trường Đại học Bách Khoa – Đại học Đà Nẵng)'],
    ['Đề tài của nhóm', 'Bảo quản hoa quả (đối tượng nghiên cứu: người bán hoa quả)'],
    ['Thời gian', '08h00 – 09h30, thứ Tư, ngày 30/09/2026'],
    ['Hình thức', 'Họp nhóm trực tiếp, kết hợp trao đổi trên nhóm Zalo'],
    ['Thành viên tham dự', 'Tuyết Ngân, Thanh An, Tuấn Anh, Gia Huy, Lê Đạt (đủ 5/5 thành viên)'],
    ['Điều phối và ghi chép', 'Tuyết Ngân'],
    ['Vị trí trong tiến trình', 'Sau khi đã phỏng vấn, tổng hợp nội dung phỏng vấn và làm Empathy Map; trước khi thực hiện Data Points & Theme Clustering, Persona & POV, HMW.'],
  ]));
  c.push(spacer());

  // 1
  c.push(sec('1. MỤC TIÊU CUỘC HỌP'));
  c.push(para('Bước Xác định vấn đề (Define) gồm nhiều công cụ nối tiếp nhau. Nếu mỗi thành viên hiểu và làm theo một cách thì kết quả cuối sẽ không khớp nhau và dễ đi lệch phương pháp. Vì vậy cuộc họp được tổ chức trước khi bắt tay vào làm, nhằm:'));
  c.push(bullet('Thống nhất cách hiểu và trình tự thực hiện bước Xác định vấn đề theo hướng dẫn của giảng viên: Data Points & Theme Clustering → Persona & POV → HMW.'));
  c.push(bullet('Rà soát dữ liệu thấu cảm đã thu thập (phỏng vấn, Empathy Map) để xác định hướng rút Theme, Insight và chọn Persona.'));
  c.push(bullet('Thống nhất hướng trả lời hai câu hỏi mà bài tập tuần yêu cầu: POV và “How might we” (HMW).'));
  c.push(bullet('Phân công hoàn thiện sản phẩm và báo cáo sau cuộc họp.'));

  // 2
  c.push(sec('2. CĂN CỨ VÀ TÀI LIỆU SỬ DỤNG'));
  c.push(bullet('**Yêu cầu bài tập tuần 3** (thông báo trên nhóm lớp): hoàn thành hoạt động nhóm; hoàn thiện báo cáo Word đến hết bước “Xác định vấn đề”; sau khi phỏng vấn thì họp và rà soát lại vấn đề để trả lời được câu hỏi HMW và POV; sản phẩm nộp gồm báo cáo Word và nhật ký nhóm.'));
  c.push(bullet('**Slide bài giảng mục 2.2 “Xác định vấn đề – Các công cụ thực hiện”** (slide 28): Data Points & Theme Clustering; Tạo dựng Persona & Bản mẫu POV; Câu hỏi kiến tạo HMW.'));
  c.push(bullet('**Cẩm nang & khung mẫu Bước 2.2 – Xác định vấn đề** do Lê Đạt tìm hiểu và phổ biến trước buổi họp (quy tắc, ví dụ đúng – sai, câu hỏi tự kiểm tra, checklist cho từng công cụ).'));
  c.push(bullet('**Dữ liệu thấu cảm của nhóm:** video và bản tổng hợp nội dung ba cuộc phỏng vấn người bán hoa quả (Tuấn Anh tổng hợp), bản phỏng vấn cô Hiền (chủ vườn bơ và sầu riêng), Empathy Map (Gia Huy).'));
  c.push(bullet('**Khảo sát Google Form** “Khảo sát về việc bảo quản hoa quả” (dữ liệu tham khảo).'));
  c.push(spacer());

  // 3
  c.push(sec('3. TÌNH HÌNH TRƯỚC CUỘC HỌP'));
  c.push(para('Tuyết Ngân báo cáo tiến độ các đầu việc đã hoàn thành làm đầu vào cho cuộc họp:'));
  c.push(table([3700, 2000, 3371], ['Công việc', 'Người thực hiện', 'Tình trạng'], [
    ['Lên kịch bản và thực hiện phỏng vấn', 'Thanh An', 'Hoàn thành (29/09)'],
    ['Quay video phỏng vấn', 'Tuấn Anh', 'Hoàn thành (29/09)'],
    ['Edit video', 'Lê Đạt', 'Hoàn thành (29/09)'],
    ['Tổng hợp nội dung phỏng vấn', 'Tuấn Anh', 'Hoàn thành (29/09)'],
    ['Empathy Map', 'Gia Huy', 'Hoàn thành (29/09)'],
    ['Tìm hiểu trước các bước Xác định vấn đề', 'Lê Đạt', 'Hoàn thành (29/09)'],
    ['Data Points & Theme Clustering; Persona & POV; HMW', 'Gia Huy; Lê Đạt', 'Chưa thực hiện – nội dung cần thống nhất trong cuộc họp'],
  ], { center: [1] }));
  c.push(spacer());

  // 4
  c.push(sec('4. DIỄN BIẾN THẢO LUẬN VÀ KẾT LUẬN'));

  // 4.1
  c.push(sub('4.1. Thống nhất cách hiểu về bước Xác định vấn đề'));
  c.push(say('Tuyết Ngân', 'Hôm nay mình họp để chốt cách làm trước khi bắt tay vào Data Points, POV và HMW. Theo yêu cầu tuần này, sau khi phỏng vấn mình phải rà soát lại vấn đề và trả lời được POV và HMW. Trước hết, mỗi bạn hiểu bước Define là làm gì?'));
  c.push(say('Gia Huy', 'Ban đầu mình nghĩ bước này là chọn luôn giải pháp cho người dùng.'));
  c.push(say('Thanh An', 'Mình cũng nghĩ vậy. Nhóm đã có hướng chế phẩm sinh học từ lúc phỏng vấn nên mình sợ khi viết POV sẽ lỡ nhắc luôn đến sản phẩm.'));
  c.push(say('Tuyết Ngân', 'Mình mở lại slide 28 của cô. Cô chỉ nêu ba công cụ là Data Points & Theme Clustering, Persona & POV, và HMW. Cả ba đều nằm ở khâu xác định vấn đề chứ chưa phải giải pháp.'));
  c.push(say('Lê Đạt', 'Mình có đọc khung mẫu. Define nằm giữa Empathize và Ideate, nhiệm vụ là biến dữ liệu rời rạc về người dùng thành một vấn đề rõ ràng, có trọng tâm, chưa đưa giải pháp. Khung mẫu có năm nguyên tắc: lấy người dùng làm trung tâm, chưa đưa giải pháp, cụ thể, truy vết được, và làm việc nhóm, phản biện thành tiếng. Mình cũng đối chiếu các bước tổng quát về xác định vấn đề với ba công cụ của cô.'));
  c.push(para('Lê Đạt trình bày bảng đối chiếu, cả nhóm cùng rà soát:', { before: 60 }));
  c.push(table([3500, 5571], ['Bước tổng quát', 'Tương ứng trong bài của nhóm'], [
    ['Quan sát thực tế; thu thập thông tin từ người dùng', 'Đã thực hiện ở bước Thấu cảm: POEMS, phỏng vấn, Empathy Map'],
    ['Tìm hiểu nguyên nhân', 'Rút Insight bằng cách hỏi “Vì sao?” từ các Theme'],
    ['Phân tích các phương pháp hiện có', 'Thể hiện trong Theme về cách người bán đang xử lý hoa quả'],
    ['Xác định khoảng trống và phát biểu vấn đề', 'Persona và POV'],
    ['Đặt tiêu chí giải pháp; kiểm chứng vấn đề', 'Thuộc giai đoạn sau, ngoài phạm vi tuần 3'],
  ]));
  c.push(spacer());
  c.push(say('Thanh An', 'Vậy hướng chế phẩm sinh học vẫn giữ, nhưng để dành cho bước lên ý tưởng đúng không?'));
  c.push(say('Lê Đạt', 'Đúng. Data point, Theme, POV, HMW đều không được nhắc đến sản phẩm hay thiết bị cụ thể, kể cả hướng nhóm đang nghiên cứu.'));
  c.push(concl([
    'Define là làm rõ **ai đang gặp vấn đề gì và vì sao**, chưa đưa ra giải pháp. Hướng chế phẩm sinh học của nhóm không xuất hiện trong Data Points, Theme, POV, HMW; nó được để lại cho bước Lên ý tưởng.',
    'Thực hiện theo thứ tự: Data Points & Theme Clustering → Persona & POV → HMW; sản phẩm trước làm đầu vào cho sản phẩm sau.',
    'Mọi kết quả phải truy vết được: HMW → POV → Theme → Data Point → dữ liệu phỏng vấn.',
  ]));
  c.push(spacer());

  // 4.2
  c.push(sub('4.2. Rà soát dữ liệu đầu vào'));
  c.push(say('Tuấn Anh', 'Mình tóm tắt trước ba cuộc phỏng vấn người bán hoa quả. Hoa quả thường bày bán khoảng 2–4 ngày rồi thay, một số loại dùng bán và làm nước thì để được 3–4 ngày. Người bán xử lý bằng tủ lạnh hoặc chỉ mua vừa đủ bán trong ngày, còn khi bày bán thì gần như không bảo quản thêm.'));
  c.push(say('Thanh An', 'Phần cô Hiền hơi khác vì cô là người trồng. Cô làm nghề khoảng 7 năm, chủ yếu có bơ và sầu riêng. Sầu riêng chín rụng trên cây giữ được tầm 5–7 ngày và cần nhiệt độ từ 25 độ trở xuống. Cô chưa có kinh nghiệm bảo quản và chưa dùng hóa chất vì lo cho sức khỏe người dùng, nên lượng bỏ đi khá nhiều và phải bán rẻ, bán xả. Cô nói nếu giúp đỡ thiệt hại kinh tế thì cô sẽ nghiên cứu để sử dụng.'));
  c.push(say('Gia Huy', 'Mình đã đưa các ý này vào Empathy Map. Ở ô Say và Do, người bán nói và làm theo cách quen thuộc. Ở ô Think và Feel, họ thấy cách hiện tại dễ làm và đủ dùng, ngại sản phẩm mới nếu không biết dùng hoặc mất thêm thời gian; những người gặp vấn đề rõ hơn về độ tươi thì cởi mở hơn.'));
  c.push(say('Lê Đạt', 'Mình thấy cùng là người bán nhưng có hai kiểu. Có người thấy cách hiện tại là đủ, có người đã mong giữ hoa quả tươi lâu hơn. Đây không phải mâu thuẫn trong cùng một người mà là khác biệt giữa hai nhóm, nên sau này không thể gộp thành một persona.'));
  c.push(say('Thanh An', 'Ô Think và Feel là phần nhóm suy ra từ lời nói. Khi viết data point mình chỉ lấy điều người ta nói và làm, đừng chép phần suy luận.'));
  c.push(say('Tuyết Ngân', 'Mình cũng đã mở Google Form khảo sát về bảo quản hoa quả. Đến lúc họp mới có 4 phản hồi, đều là học sinh, sinh viên dưới 25 tuổi, tức là người mua chứ không phải người bán. Mình đề xuất chỉ dùng để tham khảo, còn Data Points lấy từ ba cuộc phỏng vấn để dễ truy vết.'));
  c.push(para('Cả nhóm đồng ý.', { indent: { left: 284 } }));
  c.push(concl([
    'Nguồn chính của bước Define: ba cuộc phỏng vấn (gồm phỏng vấn cô Hiền) và Empathy Map. Khảo sát Google Form chỉ để tham khảo.',
    'Các ô Think, Feel của Empathy Map dùng để định hướng Insight, không chép nguyên thành data point.',
    'Ghi nhận người bán không đồng nhất về nhu cầu bảo quản; vấn đề này được xử lý ở bước Theme và Persona.',
  ]));
  c.push(spacer());

  // 4.3
  c.push(sub('4.3. Thảo luận về Data Points'));
  c.push(say('Tuyết Ngân', 'Theo mọi người, Data Point là gì?'));
  c.push(say('Tuấn Anh', 'Em nghĩ là những thông tin quan trọng mình thu được từ người được phỏng vấn.'));
  c.push(say('Lê Đạt', 'Khung mẫu ghi thêm là mẩu thông tin nhỏ nhất và chỉ chứa một ý.'));
  c.push(say('Gia Huy', 'Có người bán nói hoa quả dùng được 2–4 ngày và họ để trong tủ lạnh. Mình ghi chung thành một data point được không?'));
  c.push(say('Lê Đạt', 'Không được. Một ý nói về thời gian sử dụng, một ý nói về cách bảo quản.'));
  c.push(say('Thanh An', 'Ghi chung thì sau này gom nhóm sẽ bị lẫn.'));
  c.push(say('Tuyết Ngân', 'Vậy câu “không muốn dùng cách mới vì không biết cách dùng và mất thêm thời gian” thì sao?'));
  c.push(say('Tuấn Anh', 'Tách thành hai: một ý là không biết cách sử dụng, một ý là cho rằng sẽ tốn thời gian.'));
  c.push(para('Từ các ví dụ trên, nhóm thống nhất nguyên tắc viết Data Point và chia dữ liệu thành năm nhóm thông tin ban đầu để dễ rà soát.', { before: 60 }));
  c.push(concl([
    'Mỗi Data Point chỉ ghi **một ý**; câu trả lời chứa nhiều ý thì tách ra.',
    'Chỉ ghi những gì xuất hiện trong dữ liệu phỏng vấn, không tự diễn giải hay suy luận thêm; không chứa giải pháp.',
    'Năm nhóm thông tin: (1) thời gian sử dụng hoa quả; (2) cách người bán đang xử lý hoa quả; (3) nhận thức về nhu cầu bảo quản; (4) rào cản khi thay đổi cách làm; (5) nhu cầu thực tế trong công việc.',
  ]));
  c.push(spacer());

  // 4.4
  c.push(sub('4.4. Thảo luận về Theme Clustering'));
  c.push(para('Nhóm đặt các data point cạnh nhau, ghép những ý nói về cùng một chuyện trước, sau đó mới thảo luận để đặt tên cho từng cụm.'));
  c.push(say('Thanh An', 'Hay mình đặt tên Theme là “Bảo quản hoa quả”.'));
  c.push(say('Gia Huy', 'Mình cũng thấy ổn.'));
  c.push(say('Lê Đạt', 'Đặt vậy thì giống tên chủ đề hơn là kết luận rút ra từ dữ liệu.'));
  c.push(say('Tuyết Ngân', 'Khung mẫu yêu cầu tên Theme là một câu nhận định, nói rõ người dùng đang gặp hoặc làm điều gì. Thử che các data point đi, chỉ đọc tên Theme mà vẫn hiểu người dùng gặp chuyện gì thì mới đạt.'));
  c.push(para('Nhóm xem lại từng cụm:', { before: 60 }));
  c.push(say('Tuấn Anh', 'Cụm thời gian: 2–4 ngày, 3–4 ngày, tùy từng loại, phải thay sau một thời gian. Điểm chung là thời gian sử dụng luôn bị giới hạn.'));
  c.push(say('Gia Huy', 'Cụm cách xử lý: mình đề xuất Theme là “Sử dụng tủ lạnh để bảo quản”.'));
  c.push(say('Thanh An', 'Không phải ai cũng dùng tủ lạnh. Có người chỉ mua đủ bán trong ngày, có người không bảo quản gì.'));
  c.push(say('Tuyết Ngân', 'Vậy điểm chung thực sự là gì?'));
  c.push(say('Lê Đạt', 'Họ đều đang dùng cách sẵn có và quen thuộc với mình.'));
  c.push(say('Lê Đạt', 'Cụm nhận thức: có người nói hoa quả Việt Nam thường không cần bảo quản, có người muốn giữ tươi lâu hơn. Nhu cầu của người bán không giống nhau.'));
  c.push(say('Tuấn Anh', 'Cụm rào cản: không biết cách dùng, sợ tốn thời gian, không muốn tìm hiểu thêm. Điểm chung là chưa quen với cách mới nên khó đổi.'));
  c.push(say('Gia Huy', 'Data point “chỉ mua đủ số lượng bán trong ngày” xuất hiện ở cả ba cụm đầu. Mình để ở đâu?'));
  c.push(say('Lê Đạt', 'Có thể để ở nhiều Theme nếu mỗi Theme dùng một khía cạnh khác: ở Theme 1 là hệ quả của thời gian sử dụng giới hạn, ở Theme 2 là cách xử lý, ở Theme 3 là bằng chứng không phải ai cũng cần bảo quản thêm.'));
  c.push(para('Cả nhóm thống nhất 4 Theme. Về mức ưu tiên: Theme 1 và Theme 3 liên quan trực tiếp đến vấn đề và nhu cầu của người dùng nên được ưu tiên; Theme 2 và Theme 4 giải thích cách làm hiện tại và rào cản thay đổi, dùng làm bối cảnh cho Persona và POV.', { before: 60 }));
  c.push(concl(['Bốn Theme được thống nhất (tên Theme là câu nhận định, truy được về data point):']));
  c.push(table([1000, 4000, 4071], ['Theme', 'Tên Theme', 'Điểm chung rút ra từ data point'], [
    ['T1', 'Thời gian sử dụng hoa quả có giới hạn', 'Hoa quả chỉ dùng được vài ngày, tùy loại, và phải thay mới sau một khoảng thời gian'],
    ['T2', 'Người bán đang dựa vào cách bảo quản quen thuộc', 'Tủ lạnh, mua đủ trong ngày, không bảo quản thêm khi bày bán'],
    ['T3', 'Nhu cầu bảo quản không giống nhau giữa những người bán', 'Có người cho là không cần, có người muốn giữ tươi lâu hơn và sẵn sàng xem xét cách khác'],
    ['T4', 'Sự không quen thuộc khiến người bán khó thay đổi cách làm', 'Không biết cách dùng, sợ tốn thời gian, không muốn tìm hiểu thêm, tiếp tục cách cũ'],
  ], { center: [0] }));
  c.push(spacer());

  // 4.5
  c.push(sub('4.5. Rút Insight bằng câu hỏi “Vì sao?”'));
  c.push(para('Phần này nhóm mất nhiều thời gian nhất vì nhiều lúc mọi người chỉ mô tả lại dữ liệu thay vì giải thích nguyên nhân.'));
  c.push(say('Tuấn Anh', 'Mình đề xuất Insight là người bán muốn bảo quản hoa quả lâu hơn.'));
  c.push(say('Tuyết Ngân', 'Câu đó mới là nhu cầu thôi.'));
  c.push(say('Lê Đạt', 'Insight phải trả lời được “vì sao”, và tốt nhất là điều không quá hiển nhiên. Mình hỏi “Vì sao?” liên tục và mỗi câu trả lời phải dẫn về một theme.'));
  c.push(table([3000, 4271, 1800], ['Câu hỏi', 'Câu trả lời rút từ dữ liệu', 'Căn cứ'], [
    ['Vì sao người bán phải để ý thời gian sử dụng?', 'Hoa quả chỉ dùng được khoảng 2–4 ngày, tùy loại, và phải thay sau một khoảng thời gian.', 'Theme 1'],
    ['Vì sao nhiều người chỉ mua đủ trong ngày hoặc dùng tủ lạnh?', 'Đó là cách sẵn có, quen thuộc để hoa quả không bị để quá lâu; khi bày bán thì không bảo quản thêm.', 'Theme 2'],
    ['Vì sao nhiều người chưa muốn đổi cách làm?', 'Cách hiện tại vẫn đáp ứng việc bán hàng; cách mới bị cho là khó dùng hoặc tốn thời gian.', 'Theme 3, 4'],
    ['Vì sao một số người lại sẵn sàng xem xét cách khác?', 'Họ chịu ảnh hưởng trực tiếp: phải dùng hoa quả nhiều ngày, hoặc chịu thiệt hại khi hoa quả hỏng (như phải bán rẻ, bỏ đi ở phỏng vấn cô Hiền).', 'Theme 3; phỏng vấn cô Hiền'],
  ]));
  c.push(spacer());
  c.push(concl([
    'Insight được thống nhất: người bán chủ yếu dựa vào những cách xử lý quen thuộc để kiểm soát thời gian sử dụng hoa quả vì cách đó phù hợp với công việc hằng ngày. Họ chỉ có nhu cầu rõ ràng và sẵn sàng cân nhắc cách làm khác khi thời gian sử dụng giới hạn ảnh hưởng trực tiếp đến việc mua, dự trữ và sử dụng nguyên liệu, với điều kiện cách đó không làm công việc phức tạp hơn.',
    'Insight này giải thích được cả bốn Theme và dữ liệu từ các cuộc phỏng vấn.',
  ]));
  c.push(spacer());

  // 4.6
  c.push(sub('4.6. Chọn Persona'));
  c.push(say('Tuyết Ngân', 'Nếu phải chọn một nhóm người dùng đại diện thì mình chọn ai?'));
  c.push(say('Gia Huy', 'Từ Empathy Map và các Theme, mình thấy có hai Persona. Persona 1 là người bán duy trì cách làm quen thuộc. Persona 2 là người bán nước trái cây và hoa quả có nhu cầu giữ hoa quả tươi lâu hơn.'));
  c.push(say('Lê Đạt', 'Persona phải xây từ dữ liệu. Chi tiết nào chưa có dữ liệu trực tiếp thì đánh dấu là giả định. Mình đề xuất chọn theo ba tiêu chí: có vấn đề thực tế, có nhu cầu rõ ràng, có sự sẵn sàng thay đổi.'));
  c.push(para('Cả nhóm cùng đối chiếu:', { before: 60 }));
  c.push(table([2300, 3385, 3386], ['Tiêu chí', 'Persona 1 – duy trì cách làm quen thuộc', 'Persona 2 – cần giữ hoa quả tươi lâu hơn'], [
    ['Vấn đề thực tế', 'Chưa rõ, vì cách hiện tại vẫn bán được', 'Có: thời gian sử dụng giới hạn ảnh hưởng đến việc mua và dùng nguyên liệu'],
    ['Nhu cầu rõ ràng', 'Duy trì hoa quả đủ để bán, không làm việc phức tạp hơn', 'Giữ hoa quả tươi lâu hơn, chủ động hơn khi mua và sử dụng'],
    ['Sẵn sàng thay đổi', 'Thấp, chưa thấy lợi ích rõ ràng', 'Cao hơn: sẽ xem xét, mua nếu có cách phù hợp'],
  ]));
  c.push(spacer());
  c.push(say('Thanh An', 'Persona 2 đủ cả ba tiêu chí. Persona 1 mình vẫn giữ để thấy rõ rào cản thay đổi.'));
  c.push(concl([
    'Chọn **Persona 2 – người bán nước trái cây và hoa quả có nhu cầu giữ hoa quả tươi lâu hơn** làm người dùng mục tiêu.',
    'Giữ Persona 1 làm đối chiếu về hành vi và rào cản; chi tiết chưa có dữ liệu trực tiếp được đánh dấu là giả định.',
  ]));
  c.push(spacer());

  // 4.7
  c.push(sub('4.7. Xây dựng POV'));
  c.push(say('Tuyết Ngân', 'Công thức cô hướng dẫn là: Người dùng [Persona] + cần [Nhu cầu] + vì [Insight].'));
  c.push(say('Gia Huy', 'Mình nháp thử: “Người bán cần một sản phẩm bảo quản dễ dùng vì họ ngại thay đổi”.'));
  c.push(say('Tuyết Ngân', 'Câu này có “sản phẩm” nên nhu cầu đã là giải pháp. Nhu cầu phải bắt đầu bằng động từ.'));
  c.push(say('Lê Đạt', 'Hỏi lại là họ muốn làm được gì: duy trì độ tươi của hoa quả trong thời gian phù hợp với việc bán hàng.'));
  c.push(say('Thanh An', 'Phần “vì” cũng phải lấy từ Insight, không chỉ nhắc lại nhu cầu: vì thời gian sử dụng hoa quả có giới hạn và ảnh hưởng đến việc mua, dự trữ và sử dụng nguyên liệu.'));
  c.push(para('Nhóm ghép lại và đọc thành tiếng cho cả nhóm phản biện. POV chính được thống nhất:', { before: 60 }));
  c.push(concl(['**POV 1:** Người bán nước trái cây và hoa quả cần có khả năng duy trì độ tươi của hoa quả trong thời gian phù hợp với hoạt động bán hàng vì thời gian sử dụng của hoa quả có giới hạn và ảnh hưởng đến việc mua, dự trữ và sử dụng nguyên liệu.'], 'POV chính'));
  c.push(spacer());
  c.push(para('Để đủ 2–4 POV theo khung mẫu, nhóm viết thêm ba POV ở các góc nhìn khác, cùng đúng công thức và không chứa giải pháp:'));
  c.push(table([1000, 4300, 3771], ['POV', 'Góc tập trung', 'Xuất phát từ'], [
    ['POV 1', 'Thời gian sử dụng của hoa quả', 'Theme 1; Persona 2'],
    ['POV 2', 'Sự chủ động trong việc mua và sử dụng hoa quả', 'Theme 1 + Theme 3'],
    ['POV 3', 'Khoảng cách giữa nhu cầu và cách xử lý hiện tại', 'Theme 1 + Theme 3'],
    ['POV 4', 'Hành vi thay đổi khi nhu cầu xuất hiện', 'Theme 3 + Theme 4'],
  ], { center: [0] }));
  c.push(spacer());
  c.push(concl([
    'Giữ 4 POV, trong đó POV 1 là POV chính; nội dung đầy đủ của từng POV do Gia Huy hoàn thiện trong báo cáo, kèm phần truy vết về Theme.',
    'Kiểm tra theo khung mẫu: nhu cầu bắt đầu bằng động từ; không nhắc sản phẩm hay công nghệ; từ POV vẫn có thể nghĩ ra nhiều hướng giải pháp khác nhau.',
  ]));
  c.push(spacer());

  // 4.8
  c.push(sub('4.8. Chuyển POV thành câu hỏi HMW'));
  c.push(say('Lê Đạt', 'Mình bóc tách POV thành nhu cầu, rào cản và bối cảnh. Nhu cầu là duy trì độ tươi; rào cản là cách mới bị cho là tốn thời gian, chưa quen; bối cảnh là phải mua nhiều lần hoặc phải bán rẻ khi hoa quả hỏng. Từ mỗi yếu tố mình viết nhiều câu bắt đầu bằng “Làm thế nào để chúng ta…”.'));
  c.push(say('Tuấn Anh', 'Mình nháp: “Làm thế nào để chúng ta tạo ra thiết bị bảo quản tốt hơn?”'));
  c.push(say('Tuyết Ngân', 'Câu này nhảy sang giải pháp rồi, lại còn quá rộng vì “tốt hơn” không gắn với Insight nào.'));
  c.push(say('Lê Đạt', 'Mình bỏ “thiết bị”, thay bằng điều người bán muốn đạt được. Mỗi câu cũng chỉ nên có một mục tiêu, không dùng “và” để gộp nhiều việc.'));
  c.push(para('Sau khi điều chỉnh, nhóm thống nhất một số HMW định hướng, mỗi câu xuất phát từ một POV:', { before: 60 }));
  c.push(table([600, 1000, 7471], ['#', 'Từ POV', 'HMW định hướng'], [
    ['1', 'POV 1', 'Làm thế nào để chúng ta giúp người bán duy trì độ tươi của hoa quả trong thời gian phù hợp với hoạt động bán hàng?'],
    ['2', 'POV 1', 'Làm thế nào để chúng ta giúp người bán giữ được giá trị của hoa quả cho đến khi bán hết?'],
    ['3', 'POV 2', 'Làm thế nào để chúng ta giúp người bán chủ động hơn về lượng hoa quả cần mua trong mỗi lần nhập?'],
    ['4', 'POV 3', 'Làm thế nào để chúng ta giúp người bán có thêm thời gian sử dụng hoa quả mà vẫn giữ được cách bán hằng ngày?'],
    ['5', 'POV 4', 'Làm thế nào để chúng ta giúp người bán thử một cách làm mới mà công việc hằng ngày không phức tạp hơn?'],
  ], { center: [0, 1] }));
  c.push(spacer());
  c.push(concl([
    'HMW là câu hỏi mở, bắt đầu bằng “Làm thế nào để chúng ta…”, giữ tinh thần tích cực, không chứa giải pháp, mỗi câu một mục tiêu, không quá rộng cũng không quá hẹp.',
    'Lê Đạt mở rộng từ các câu định hướng trên thành 5–10 HMW, tự kiểm tra bằng bảng chấm điểm của khung mẫu, chọn những câu tốt nhất và bảo đảm phủ đều các POV.',
  ]));
  c.push(spacer());

  // 5
  c.push(sec('5. TỔNG HỢP KẾT LUẬN CỦA CUỘC HỌP'));
  c.push(table([2300, 6771], ['Nội dung', 'Kết luận thống nhất'], [
    ['Cách hiểu bước Define', 'Làm rõ người dùng gặp vấn đề gì và vì sao; chưa đưa ra giải pháp hay sản phẩm.'],
    ['Nguồn dữ liệu', 'Ba cuộc phỏng vấn (gồm cô Hiền) và Empathy Map; khảo sát Google Form chỉ để tham khảo.'],
    ['Data Points', 'Mỗi data point một ý, không suy luận, không chứa giải pháp; năm nhóm thông tin.'],
    ['Theme', '4 Theme dạng câu nhận định; ưu tiên Theme 1 và Theme 3.'],
    ['Insight', 'Người bán dựa vào cách quen thuộc; chỉ cân nhắc đổi khi thời gian sử dụng giới hạn ảnh hưởng trực tiếp đến việc mua, dự trữ, sử dụng hoặc gây thiệt hại.'],
    ['Persona', 'Chọn Persona 2 – người bán nước trái cây và hoa quả có nhu cầu giữ hoa quả tươi lâu hơn.'],
    ['POV', '4 POV, POV 1 là POV chính.'],
    ['HMW', '5–10 câu mở, không chứa giải pháp, truy vết được về POV.'],
  ], { labelCol: true }));
  c.push(spacer());

  // 6
  c.push(sec('6. KHÓ KHĂN VÀ BÀI HỌC RÚT RA'));
  c.push(bullet('**Nhầm Define với tìm giải pháp:** nhiều thành viên ban đầu nghĩ phải chọn sản phẩm. Sau khi đối chiếu với slide và khung mẫu, nhóm thống nhất giữ giải pháp lại cho bước Lên ý tưởng.'));
  c.push(bullet('**Đặt tên Theme như tên chủ đề:** ví dụ “Bảo quản hoa quả”, “Sử dụng tủ lạnh”. Cách sửa là hỏi điểm chung của cụm và viết thành câu nhận định.'));
  c.push(bullet('**Nhầm nhu cầu với giải pháp, insight với nhu cầu:** các bản nháp POV và HMW ban đầu có “sản phẩm”, “thiết bị”; nhóm sửa bằng cách hỏi “họ muốn làm được gì?” và “vì sao?”.'));
  c.push(bullet('**Dữ liệu chưa đồng nhất giữa các người bán:** nhóm nhận ra cần tách hai Persona thay vì gộp một.'));
  c.push(spacer());

  // 7
  c.push(sec('7. PHÂN CÔNG SAU CUỘC HỌP'));
  c.push(table([2700, 1500, 1500, 3371], ['Công việc', 'Người phụ trách', 'Thời hạn (30/09)', 'Yêu cầu sản phẩm'], [
    ['Data Points & Theme Clustering', 'Gia Huy', '15h00', 'Data Points theo năm nhóm, 4 Theme, insight nháp, bản đồ truy vết Data Point → Theme → Insight'],
    ['Persona & POV', 'Gia Huy', '17h30', '2 Persona, chọn Persona mục tiêu và lý do; 4 POV có truy vết'],
    ['HMW', 'Lê Đạt', '18h45', '5–10 HMW chọn lọc, truy vết được về POV'],
    ['Rà soát, đối chiếu phương pháp và dữ liệu', 'Thanh An, Tuấn Anh', 'Trước 23h30', 'Thanh An đối chiếu phương pháp và checklist; Tuấn Anh đối chiếu với bản tổng hợp phỏng vấn gốc'],
    ['Nhật ký nhóm (nhật ký làm việc và nhật ký họp)', 'Tuyết Ngân', '22h30', 'Nhật ký tuần 3 theo yêu cầu nộp'],
    ['Báo cáo Word đến hết bước Xác định vấn đề', 'Tuyết Ngân', '23h30', 'Báo cáo tổng hợp theo yêu cầu của giảng viên'],
  ], { center: [1, 2] }));
  c.push(spacer());

  // 8
  c.push(sec('8. KẾT THÚC CUỘC HỌP'));
  c.push(say('Tuyết Ngân', 'Điều quan trọng nhất hôm nay là cả nhóm đã hiểu rõ hơn cách làm bước Define và có chung một hướng, không phải là xong ngay sản phẩm. Từ giờ mọi người làm theo đúng những điểm đã thống nhất, có gì chưa chắc thì hỏi lại trên nhóm.'));
  c.push(para('Cuộc họp kết thúc lúc 09h30 cùng ngày. Các thành viên tham dự đã đọc và thống nhất nội dung nhật ký này.', { before: 60 }));
  c.push(para('Ngày 30 tháng 9 năm 2026', { align: AlignmentType.RIGHT, before: 120, after: 0 }));
  c.push(para('**Người điều phối và ghi chép**', { align: AlignmentType.RIGHT, after: 600 }));
  c.push(para('**Tuyết Ngân**', { align: AlignmentType.RIGHT }));
  return c;
}

// ---------- Nhật ký làm việc ----------
function workLog() {
  const c = [];
  c.push(info([
    ['Học phần', 'KNM – Hoạt động nhóm tuần 3'],
    ['Giảng viên', 'TS. Lê Lý Thùy Trâm (Bộ môn Công nghệ sinh học, Trường Đại học Bách Khoa – Đại học Đà Nẵng)'],
    ['Đề tài của nhóm', 'Bảo quản hoa quả (đối tượng nghiên cứu: người bán hoa quả)'],
    ['Nội dung của tuần', 'Bước 2.2 – Xác định vấn đề (Define)'],
    ['Thời gian ghi nhận', 'Từ 28/09/2026 đến 30/09/2026'],
    ['Thành viên', 'Tuyết Ngân, Thanh An, Tuấn Anh, Gia Huy, Lê Đạt'],
    ['Sản phẩm nộp', 'Báo cáo Word (đến hết bước Xác định vấn đề) và Nhật ký nhóm'],
  ]));
  c.push(spacer());

  c.push(h('PHẦN I. NHẬT KÝ LÀM VIỆC', 1));
  c.push(h('1. Mục tiêu và yêu cầu của tuần', 2));
  c.push(para('Theo thông báo bài tập tuần 3 trên nhóm lớp, nhóm cần thực hiện:'));
  c.push(bullet('Tiếp tục hoàn thành hoạt động nhóm và có báo cáo nhóm.'));
  c.push(bullet('Hoàn thiện báo cáo Word đến hết bước “Xác định vấn đề”.'));
  c.push(bullet('Thực hiện phỏng vấn, sau đó họp và rà soát lại vấn đề để trả lời được câu hỏi “How might we” (HMW) và POV.'));
  c.push(para('Sản phẩm nộp: báo cáo Word và nhật ký nhóm. Theo slide mục 2.2, bước Xác định vấn đề gồm ba công cụ: Data Points & Theme Clustering; Persona & Bản mẫu POV; Câu hỏi kiến tạo HMW.', { before: 60 }));

  c.push(h('2. Bảng phân công và tiến độ công việc', 2));
  c.push(para('Các đầu việc được giao ngày 28/09/2026 và theo dõi trên bảng phân công của nhóm:'));
  const D = (d, t) => `${d}\n${t}`;
  c.push(table([450, 2050, 1150, 1250, 1050, 3121], ['TT', 'Việc cần làm', 'Người phụ trách', 'Hạn hoàn thành', 'Trạng thái', 'Sản phẩm / Ghi chú'], [
    ['1', 'Làm POEMS', 'Thanh An', D('28/09', '19h00'), 'Hoàn thành', 'Bảng quan sát phục vụ bước Thấu cảm'],
    ['2', 'Lên kịch bản phỏng vấn', 'Thanh An', D('28/09', '22h00'), 'Hoàn thành', 'Kịch bản gồm hai phần: thực tế bảo quản trái cây; mong muốn về giải pháp mới'],
    ['3', 'Phỏng vấn', 'Thanh An', D('29/09', '12h00'), 'Hoàn thành', 'Ba cuộc phỏng vấn người bán hoa quả; bản phỏng vấn cô Hiền (vườn bơ và sầu riêng)'],
    ['4', 'Tìm người được phỏng vấn', 'Cả nhóm', D('29/09', 'trước giờ phỏng vấn'), 'Hoàn thành', 'Liên hệ các đối tượng phù hợp'],
    ['5', 'Tổng hợp nội dung phỏng vấn', 'Tuấn Anh', D('29/09', '20h00'), 'Hoàn thành', 'Tổng hợp từ video: phỏng vấn ai, họ nói gì'],
    ['6', 'Quay video', 'Tuấn Anh', D('29/09', '14h00'), 'Hoàn thành', 'Gửi video lên nhóm để tổng hợp và chỉnh sửa'],
    ['7', 'Edit video', 'Lê Đạt', D('29/09', '23h59'), 'Hoàn thành', 'Gửi video sớm để nhóm chỉnh sửa'],
    ['8', 'Làm Empathy Map', 'Gia Huy', D('29/09', '23h59'), 'Hoàn thành', 'Empathy Map bốn ô Say – Do – Think – Feel, làm ngay khi có nội dung tổng hợp video'],
    ['9', 'Tìm hiểu trước các bước Xác định vấn đề', 'Lê Đạt', D('29/09', '22h00'), 'Hoàn thành', 'Cẩm nang & khung mẫu Bước 2.2: các bước chi tiết phải làm, đã phổ biến cho nhóm'],
    ['10', 'Họp nhóm', 'Tuyết Ngân', D('30/09', '08h00 – 09h30'), 'Hoàn thành', 'Làm rõ các bước xác định vấn đề (kết thúc trước 09h30); xem Phần II'],
    ['11', 'Data Points & Theme Clustering', 'Gia Huy', D('30/09', '15h00'), 'Hoàn thành', 'Data Points theo 5 nhóm, 4 Theme, bản đồ truy vết'],
    ['12', 'Persona & POV', 'Gia Huy', D('30/09', '17h30'), 'Hoàn thành', '2 Persona (chọn Persona 2), 4 POV có truy vết'],
    ['13', 'Câu hỏi kiến tạo HMW', 'Lê Đạt', D('30/09', '18h45'), 'Hoàn thành', 'Bộ HMW chuyển từ POV, không chứa giải pháp'],
    ['14', 'Rà soát, đối chiếu phương pháp và dữ liệu', 'Thanh An, Tuấn Anh', D('30/09', 'trước 23h30'), 'Hoàn thành', 'Đối chiếu với checklist và bản tổng hợp phỏng vấn gốc'],
    ['15', 'Nhật ký nhóm', 'Tuyết Ngân', D('30/09', '22h30'), 'Hoàn thành', 'Nhật ký làm việc và nhật ký họp nhóm'],
    ['16', 'Báo cáo bài làm (Word)', 'Tuyết Ngân', D('30/09', '23h30'), 'Hoàn thành', 'Hoàn thiện báo cáo tổng hợp theo yêu cầu'],
  ], { center: [0, 2, 3, 4] }));

  return c;
}

// ---------- Khởi tạo Document và xuất file ----------
const doc = new Document({
  numbering: {
    config: [
      {
        reference: 'bul',
        levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT }],
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          margin: { top: 1134, bottom: 1134, left: 1701, right: 1134 }, // Lề chuẩn A4: trái 3cm, phải 2cm, trên/dưới 2cm
        },
      },
      headers: {
        default: new Header({
          children: [para('KNM – Báo cáo nhật ký hoạt động nhóm tuần 3', { align: AlignmentType.RIGHT, run: { size: 18, color: '666666' } })],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 18, color: '666666' })],
            }),
          ],
        }),
      },
      children: [
        ...workLog(),
        spacer(),
        new Paragraph({ children: [new PageBreak()] }),
        h('PHẦN II. BIÊN BẢN CUỘC HỌP NHÓM', 1),
        spacer(),
        ...meeting(2),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync('Nhat_ky_lam_viec_Tuan_3.docx', buffer);
  console.log('Đã tạo thành công file Nhat_ky_lam_viec_Tuan_3.docx!');
});
