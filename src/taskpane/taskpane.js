/* Thai <-> English Kedmanee keyboard layout converter
   รองรับ Word, Excel, PowerPoint */

const engToThai = {
  '`':'_','1':'ๅ','2':'/','3':'-','4':'ภ','5':'ถ','6':'ุ','7':'ึ','8':'ค','9':'ต','0':'จ','-':'ข','=':'ช',
  '~':'๐','!':'+','@':'๑','#':'๒','$':'๓','%':'๔','^':'ู','&':'ฺ','*':'๕','(':'๖',')':'๗','_':'๘','+':'๙',
  'q':'ๆ','w':'ไ','e':'ำ','r':'พ','t':'ะ','y':'ั','u':'ี','i':'ร','o':'น','p':'ย','[':'บ',']':'ล','\\':'ฃ',
  'Q':'๐','W':'๒','E':'๓','R':'๔','T':'ธ','Y':'็','U':'ณ','I':'ญ','O':'ฐ','P':',','{':'ฎ','}':'ฑ','|':'ธ',
  'a':'ฟ','s':'ห','d':'ก','f':'ด','g':'เ','h':'้','j':'่','k':'า','l':'ส',';':'ว',"'":'ง',
  'A':'ฤ','S':'ฆ','D':'ฏ','F':'โ','G':'ฌ','H':'็','J':'๋','K':'ษ','L':'ศ',':':'ซ','"':'.',
  'z':'ผ','x':'ป','c':'แ','v':'อ','b':'ิ','n':'ื','m':'ท',',':'ม','.':'ใ','/':'ฝ',
  'Z':'(','X':')','C':'ฉ','V':'ฮ','B':'ฺ','N':'์','M':'?','<':'ฒ','>':'ฬ','?':'ฅ',
  ' ':' ','\n':'\n','\r':'\r'
};

const thaiToEng = {};
Object.entries(engToThai).forEach(([eng, thai]) => {
  if (!thaiToEng[thai]) thaiToEng[thai] = eng;
});

let currentApp = 'Word'; // default

Office.onReady((info) => {
  // detect app ที่กำลังใช้งาน
  if (info.host === Office.HostType.Excel) currentApp = 'Excel';
  else if (info.host === Office.HostType.PowerPoint) currentApp = 'PowerPoint';
  else currentApp = 'Word';

  log('✅ พร้อมใช้งานใน ' + currentApp);
});

// ---- Convert ตาม app ----
async function convertNow() {
  if (currentApp === 'Excel') {
    await convertExcel();
  } else if (currentApp === 'PowerPoint') {
    await convertPowerPoint();
  } else {
    await convertWord();
  }
}

// ---- Word ----
async function convertWord() {
  try {
    await Word.run(async (context) => {
      const sel = context.document.getSelection();
      sel.load('text');
      await context.sync();

      const original = sel.text;
      if (!original.trim()) { log('⚠️ กรุณาเลือก text ก่อน'); return; }

      const converted = convertText(original);
      if (converted === original) { log('ℹ️ ไม่พบ text ที่ต้องแปลง'); return; }

      sel.insertText(converted, 'Replace');
      await context.sync();
      log('✅ Word: แปลงสำเร็จ');
    });
  } catch (e) { log('❌ ' + e.message); }
}

// ---- Excel ----
async function convertExcel() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load('values');
      await context.sync();

      const values = range.values;
      let changed = false;

      const newValues = values.map(row => row.map(cell => {
        if (typeof cell !== 'string' || !cell.trim()) return cell;
        const converted = convertText(cell);
        if (converted !== cell) { changed = true; }
        return converted;
      }));

      if (!changed) { log('ℹ️ ไม่พบ text ที่ต้องแปลง'); return; }

      range.values = newValues;
      await context.sync();
      log('✅ Excel: แปลงสำเร็จ');
    });
  } catch (e) { log('❌ ' + e.message); }
}

// ---- PowerPoint ----
async function convertPowerPoint() {
  try {
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const shapes = slide.shapes;
      shapes.load('items');
      await context.sync();

      let changed = false;

      for (const shape of shapes.items) {
        if (shape.textFrame) {
          shape.textFrame.load('text');
        }
      }
      await context.sync();

      for (const shape of shapes.items) {
        try {
          const tf = shape.textFrame;
          if (!tf || !tf.text || !tf.text.trim()) continue;
          const converted = convertText(tf.text);
          if (converted !== tf.text) {
            tf.text = converted;
            changed = true;
          }
        } catch (e) { /* shape ไม่มี textFrame */ }
      }

      await context.sync();
      if (!changed) { log('ℹ️ ไม่พบ text ที่ต้องแปลง'); return; }
      log('✅ PowerPoint: แปลงสำเร็จ');
    });
  } catch (e) { log('❌ ' + e.message); }
}

// ---- Core conversion logic ----
function convertText(text) {
  const thaiCount = (text.match(/[\u0E00-\u0E7F]/g) || []).length;
  const engCount  = (text.match(/[a-zA-Z0-9`~!@#$%^&*\-=\[\]\\;',./{}|:"<>?_+]/g) || []).length;
  const total = text.replace(/\s/g, '').length;
  if (total === 0) return text;

  let result = '';
  if (engCount >= thaiCount) {
    for (const ch of text) result += engToThai[ch] !== undefined ? engToThai[ch] : ch;
    log('🔤 Eng → Thai');
  } else {
    for (const ch of text) result += thaiToEng[ch] !== undefined ? thaiToEng[ch] : ch;
    log('🔤 Thai → Eng');
  }
  return result;
}

function log(msg) {
  const box = document.getElementById('logBox');
  if (!box) return;
  const p = document.createElement('p');
  p.textContent = new Date().toLocaleTimeString('th-TH') + ' ' + msg;
  box.prepend(p);
  while (box.children.length > 20) box.removeChild(box.lastChild);
}