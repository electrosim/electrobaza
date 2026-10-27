import fs from 'node:fs/promises';
import path from 'node:path';
import {pathToFileURL} from 'node:url';
import {Presentation, PresentationFile} from 'file:///C:/Users/user/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs';

const root='D:/Electrobaza';
const build=path.join(root,'.pptx-build');
const out=path.join(build,'output');
await fs.mkdir(out,{recursive:true});
const skill='C:/Users/user/.codex/plugins/cache/openai-primary-runtime/presentations/26.1025.11240/skills/presentations';
const python='C:/Users/user/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe';
const photos=JSON.parse(await fs.readFile(path.join(build,'photos.json'),'utf8'));
const sources=Object.fromEntries(photos.map(p=>[p.id,p.source]));
const lecture='Материал Электробазы: 1ЭО-24, МДК 04.02, 3-й семестр. Лекция «Неисправности электрических аппаратов и способы их устранения. Испытания и наладка электрических аппаратов». Числовой расчёт исключён по согласованию.';
const p=Presentation.create({slideSize:{width:1280,height:720}});
const C={ink:'#12343B',body:'#253E44',muted:'#526A70',accent:'#B94721',bg:'#FAFCFB'};
function text(s,t,x,y,w,h,size=30,bold=false,color=C.body){
 const sh=s.shapes.add({geometry:'textbox',name:t.slice(0,50),position:{left:x,top:y,width:w,height:h},fill:'none',line:{fill:'none',width:0}});
 sh.text=t;sh.text.style={typeface:'Arial',fontSize:size,bold,color,verticalAlignment:'top',alignment:'left',insets:0,wrap:'square',autoFit:'none'};return sh;
}
function slide(title){const s=p.slides.add();s.background.fill=C.bg;text(s,title,60,45,1160,105,43,true,C.ink);text(s,String(p.slides.items.length).padStart(2,'0'),1174,670,50,25,18,false,C.muted);s.speakerNotes.text=lecture;return s;}
async function photoSlide(id,title,visible,check,question){
 const s=slide(title);const file=path.join(build,id+(id==='02'?'.png':'.jpg'));const raw=await fs.readFile(file);
 s.images.add({blob:new Uint8Array(raw),contentType:id==='02'?'image/png':'image/jpeg',alt:title,fit:'contain',position:{left:60,top:154,width:710,height:484}});
 text(s,'Что видно',810,164,400,40,29,true,C.accent);text(s,visible,810,211,400,146,28);
 text(s,'Что проверяют',810,382,400,40,29,true,C.accent);text(s,check,810,428,400,166,28);
 text(s,question,60,650,1080,40,22,false,C.muted);
 s.speakerNotes.text=lecture+'\nФото '+id+': '+sources[id]+'\nПоказать внешний признак, затем обсудить возможные причины. Причина конкретного отказа требует проверки.\n'+visible+'\n'+check;
}
const cover=p.slides.add();cover.background.fill=C.ink;
text(cover,'1ЭО-24   МДК 04.02   3-й семестр',64,68,1150,44,26,false,'#BFD6D5');
text(cover,'Неисправности\nэлектрических аппаратов\nи способы их устранения',64,176,1150,272,58,true,'#FFFFFF');
text(cover,'Испытания и наладка\nэлектрических аппаратов',64,496,1100,112,39,false,'#E1ECEA');
cover.speakerNotes.text=lecture+'\nПрезентация включает все 10 фотографий из последней согласованной подборки. Источники указаны в заметках к соответствующим слайдам.';
let s=slide('Что позволяет увидеть внешний осмотр');
text(s,'Контакты',60,166,520,45,32,true,C.accent);text(s,'Нагар, кратеры, неровности и потерю материала на рабочей поверхности.',60,223,520,112,32);
text(s,'Изоляция и корпус',680,166,520,45,32,true,C.accent);text(s,'Потемнение, оплавление, обугливание и трещины.',680,223,520,112,32);
text(s,'Зажимы и проводники',60,376,520,45,32,true,C.accent);text(s,'Повреждение жил, следы перегрева и продукты коррозии.',60,433,520,124,32);
text(s,'Катушка',680,376,520,45,32,true,C.accent);text(s,'Повреждение каркаса и изоляции обмотки. Не каждый дефект виден снаружи.',680,433,520,124,32);
text(s,'Осмотр выявляет признаки. Измерения и испытания помогают установить причину.',60,613,1120,56,27,true,C.ink);
await photoSlide('01','Подгорание контактной поверхности','На контактной накладке виден тёмный участок с неровной поверхностью.','Состояние обеих контактных поверхностей, нажатие и допустимый износ по документации.','Почему уменьшение площади надёжного соприкосновения усиливает нагрев?');
await photoSlide('10','Износ контактов в сравнении','Верхняя деталь потемнела. На нижней лучше различимы рабочие поверхности.','Износ накладок, состояние пружин и соответствие аппарата характеру нагрузки.','По каким внешним признакам эти детали отличаются?');
await photoSlide('02','Повреждение контактных деталей','На разобранных деталях видны потемнение и неровности контактных участков.','Целостность деталей, состояние креплений и возможность замены контактного узла.','Разборка аппарата сама по себе ещё не определяет причину отказа.');
await photoSlide('03','Обугливание корпуса выключателя','Корпус возле зажима сильно разрушен. Повреждение выходит за пределы поверхности контакта.','Состояние проводника и соединения, нагрузку и пригодность аппарата к дальнейшей работе.','Повреждённый неразборный выключатель заменяют.');
await photoSlide('04','Перегрев соединения в щите','Возле нижнего левого зажима правого аппарата видны почернение и повреждение изоляции.','Качество оконцевания, состояние жилы и зажима, соответствие проводника нагрузке.','Проверка включает и аппарат, и подключённый проводник.');
await photoSlide('05','Повреждение изоляции катушки','В средней части обмотки виден обугленный участок и разрушение изоляционного материала.','Напряжение питания, режим работы, свободное перемещение и прилегание магнитной системы.','Замена катушки требует устранения причины её перегрева.');
await photoSlide('06','Тепловое повреждение контактов реле','Рамка на исходной фотографии выделяет контактный узел со следами перегрева и оплавления.','Переключение контактов, допустимую нагрузку и состояние управляемой цепи.','Подпись на фотографии: «Перегретые контакты реле со следами оплавления».');
await photoSlide('07','Разрушение основания реле','Между выводами видна обугленная полость. Изоляционное основание потеряло целостность.','Также осматривают ответную колодку, зажимы и проводники. Повреждённое реле заменяют.','Пример автомобильного реле: внешний признак повреждения изоляции.');
await photoSlide('08','Повреждение силового контактора','Потемнение корпуса и повреждённые соединения заметны в нижней части аппарата.','Контактор, наконечники и кабели, а также нагрузку и условия эксплуатации.','Объём восстановления определяют по всем повреждённым элементам.');
await photoSlide('09','Коррозия проводников и загрязнение','Слева видны зелёные отложения на оголённых жилах. Поверхности аппаратов загрязнены.','Состояние жил и изоляции, герметичность шкафа и наличие влаги или агрессивной среды.','Очистка поверхности не восстанавливает разрушенную жилу.');

const tableOwners=[];
function tableSlide(title,rows){
 const s=slide(title);tableOwners.push(p.slides.items.length);
 const values=[['Признак','Возможная причина','Проверка и устранение'],...rows];
 const t=s.tables.add({rows:values.length,columns:3,left:60,top:168,width:1160,height:440,columnWidths:[292,374,494],values});
 t.borders.assign({fill:'#CBD8D6',width:1,style:'solid'});
 t.cells.block({row:0,column:0,rowCount:values.length,columnCount:3}).assign({textStyle:{typeface:'Arial',fontSize:27,color:C.body},margins:{left:14,right:14,top:12,bottom:12},anchor:'center'});
 for(let r=0;r<values.length;r++){
  t.rows[r].height=r===0?58:(382/rows.length);
  for(let c=0;c<3;c++){const cell=t.getCell(r,c);cell.fill=r===0?C.ink:(r%2?'#EFF5F3':'#FFFFFF');cell.text.style={typeface:'Arial',fontSize:r===0?25:27,color:r===0?'#FFFFFF':C.body,bold:r===0};}
 }
 text(s,'Допустимые значения и способы ремонта определяют по документации изготовителя.',60,635,1120,48,25,false,C.muted);
}
tableSlide('Неисправности выключателей и соединений',[
['Выключатель отключается','Перегрузка, короткое замыкание или неисправность','Проверить нагрузку и цепь. Устранить причину. Неисправный выключатель заменить.'],
['Не удерживается во включённом положении','Повреждение механизма или условие отключения','Проверить механизм и цепь по документации. При подтверждении отказа заменить аппарат.'],
['Нагревается зажим','Ослабление соединения, повреждение жилы или перегрузка','Проверить проводник и зажим. Восстановить соединение или заменить повреждённые элементы.']]);
tableSlide('Неисправности контакторов',[
['Не включается','Нет питания катушки, обрыв цепи или повреждение катушки','Проверить цепь управления и катушку. Восстановить цепь или заменить неисправный элемент.'],
['Гудит, якорь вибрирует','Пониженное напряжение или неисправность магнитной системы','Проверить питание, загрязнение и механическое состояние. Устранить выявленную причину.'],
['Не отключается','Сохранился сигнал, заедает механизм или сварились контакты','Проверить цепь управления и аппарат. Заменить повреждённый узел или аппарат.']]);
tableSlide('Неисправности реле и катушек',[
['Реле срабатывает, но цепь не включается','Повреждены контакты или нарушена внешняя цепь','Проверить переключение контактов и целостность цепи. Устранить обрыв или заменить реле.'],
['Катушка сильно нагревается','Неподходящее питание, режим работы или повреждение катушки','Сопоставить питание и режим с паспортом. Устранить причину, заменить повреждённую катушку.']]);
s=slide('Последовательность поиска неисправности');
const steps=[['1. Проявление отказа','Когда возникает отказ и какие команды его вызывают?'],['2. Схема и паспорт','Напряжение катушки, нагрузка, подключения и блокировки.'],['3. Осмотр и измерения','Сначала состояние аппарата и цепей при снятом напряжении.'],['4. Устранение и проверка','Восстановление цепи или замена с проверкой работы в схеме.']];
steps.forEach((a,i)=>{const y=165+i*112;text(s,a[0],60,y,390,58,31,true,C.accent);text(s,a[1],490,y,720,78,31);});
text(s,'До осмотра отключают питание, исключают ошибочное включение и проверяют отсутствие напряжения.',60,633,1120,56,24,true,C.ink);
s=slide('Испытания и наладка после ремонта');
text(s,'Испытания',60,165,540,50,34,true,C.accent);
text(s,'Механическая работа и состояние контактов.\n\nИзоляция и электрические параметры по методике изготовителя.\n\nЗащитные характеристики выключателя на испытательном оборудовании.',60,228,540,360,30);
text(s,'Наладка',680,165,540,50,34,true,C.accent);
text(s,'Проверка включения и отключения в составе схемы.\n\nПроверка блокировок и последовательности работы.\n\nНастройка предусмотренных уставок и оформление результатов.',680,228,540,360,30);
text(s,'Проверка механизма расцепления не подтверждает токовые характеристики защиты.',60,635,1120,55,27,true,C.ink);
s=slide('Вопросы для обсуждения');
['Какие повреждения можно определить при осмотре?',
'Почему замена аппарата без устранения причины может привести к повторному отказу?',
'Чем отличается повреждение катушки от неисправности цепи её питания?',
'Какие проверки необходимы после ремонта или замены?'].forEach((q,i)=>text(s,`${i+1}. ${q}`,60,170+i*110,1130,90,33,i===0,C.body));
s.speakerNotes.text=lecture+'\nОбсудить связь внешнего признака, проверяемой гипотезы и результата измерения. Источники фотографий приведены в заметках слайдов 3–12.';

for(let i=0;i<p.slides.items.length;i++){
 const blob=await p.export({slide:p.slides.items[i],format:'png',scale:1});
 await fs.writeFile(path.join(build,`slide-${String(i+1).padStart(2,'0')}.png`),new Uint8Array(await blob.arrayBuffer()));
 console.log('Rendered',i+1);
}
await fs.writeFile(path.join(build,'deck.json'),JSON.stringify(p.toProto()));
const candidate=path.join(build,'candidate.pptx');await(await PresentationFile.exportPptx(p)).save(candidate);
const {finalizePresentation}=await import(pathToFileURL(path.join(skill,'container_tools/artifact_tool_utils.mjs')).href);
const final=path.join(out,'Презентация — Неисправности электрических аппаратов.pptx');
const result=await finalizePresentation({workspaceDir:root,candidatePath:candidate,finalPath:final,pythonExecutable:python,integrityValidatorPath:path.join(skill,'container_tools/inspect_presentation_package_integrity.py'),layoutValidatorPath:path.join(skill,'container_tools/inspect_presentation_layout_geometry.py'),layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-bullet-geometry','--validate-heading-fit',...tableOwners.flatMap(x=>['--require-native-table-slide',String(x)])],requiredNativeTableOwnerSlides:tableOwners,fontPolicy:{basis:'design',families:['Arial']},verifyArtifactToolImport:true,receiptPath:path.join(build,'validation.json')});
console.log(JSON.stringify({slides:p.slides.items.length,result,final}));
