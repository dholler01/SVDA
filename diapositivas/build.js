const pptxgen = require('pptxgenjs');
const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';
pres.title = 'NUBA Bolivia – Dirección Estratégica 1';
const BERRY='6D2E46', ROSE='A26769', BLUSH='F8ECEF', INK='2B1B22', WHITE='FFFFFF', GREEN='3E7D5A', RED='B3362F', GREY='6B5B62';
const HF='Cambria', BF='Calibri';

function title(s, t, sub){
  s.background={color:WHITE};
  s.addText(t,{x:0.5,y:0.3,w:9,h:0.7,fontFace:HF,fontSize:32,bold:true,color:BERRY,margin:0,isTextBox:true});
  if(sub) s.addText(sub,{x:0.5,y:0.98,w:9,h:0.35,fontFace:BF,fontSize:14,color:GREY,margin:0,isTextBox:true});
}
let _n=1;function foot(s,n0){const n=++_n; s.addText('NUBA Bolivia · Dirección Estratégica 1',{x:0.5,y:5.25,w:6,h:0.25,fontFace:BF,fontSize:10,color:GREY,margin:0,isTextBox:true});
  s.addText(String(n),{x:9,y:5.25,w:0.5,h:0.25,fontFace:BF,fontSize:10,color:GREY,align:'right',margin:0,isTextBox:true}); }
const hdr=(t)=>({text:t,options:{bold:true,color:WHITE,fill:{color:BERRY},fontFace:BF,fontSize:12}});
const cell=(t,o={})=>({text:t,options:Object.assign({fontFace:BF,fontSize:12,color:INK},o)});

// 1 Portada
let s=pres.addSlide(); s.background={color:BERRY};
s.addShape(pres.ShapeType.ellipse,{x:6.6,y:-0.9,w:4.6,h:4.6,fill:{color:ROSE,transparency:55},line:{color:ROSE,transparency:100},objectName:'Circulo decorativo'});
s.addText('NUBA Bolivia',{x:0.6,y:1.4,w:8,h:1,fontFace:HF,fontSize:54,bold:true,color:WHITE,margin:0,isTextBox:true});
s.addText('Misión, visión, actores, entorno y fuerzas de Porter',{x:0.6,y:2.45,w:8,h:0.5,fontFace:BF,fontSize:20,color:'F3D9E0',margin:0,isTextBox:true});
s.addText('Dirección Estratégica 1 · UTEPSA · Santa Cruz de la Sierra, noviembre de 2026',{x:0.6,y:4.5,w:8.5,h:0.4,fontFace:BF,fontSize:14,color:WHITE,margin:0,isTextBox:true});
s.addNotes('Presentación de NUBA, cadena boliviana de belleza. Sucursal en estudio: Ventura Mall, Santa Cruz.');

// 2 Agenda
s=pres.addSlide(); title(s,'Qué vamos a ver');
const ag=[['1','Misión y visión','actual vs. mejorada'],['2','Grupos de interés','matriz de poder'],['3','Entorno','PESTEL y perfil estratégico'],['4','Cinco fuerzas de Porter','poder y rentabilidad'],['5','Conclusión','oportunidades y amenazas']];
ag.forEach((a,i)=>{const y=1.5+i*0.72;
 s.addShape(pres.ShapeType.ellipse,{x:0.6,y,w:0.52,h:0.52,fill:{color:BERRY},line:{color:BERRY},objectName:'Num '+a[0]});
 s.addText(a[0],{x:0.6,y,w:0.52,h:0.52,align:'center',valign:'middle',fontFace:HF,fontSize:18,bold:true,color:WHITE,margin:0,isTextBox:true});
 s.addText([{text:a[1],options:{bold:true,fontSize:20,color:INK,breakLine:true}},{text:a[2],options:{fontSize:14,color:GREY}}],{x:1.35,y:y-0.06,w:7,h:0.65,fontFace:BF,margin:0,valign:'middle',isTextBox:true});});
foot(s,2);

// 3 Misión y visión actuales
s=pres.addSlide(); title(s,'Misión y visión actuales','NUBA no publica una misión y visión formales; esto es lo más cercano');
s.addShape(pres.ShapeType.roundRect,{x:0.5,y:1.55,w:4.3,h:3.3,rectRadius:0.12,fill:{color:BLUSH},line:{color:BLUSH},objectName:'Tarjeta mision'});
s.addText('MISIÓN ACTUAL',{x:0.75,y:1.7,w:3.8,h:0.3,fontFace:BF,fontSize:13,bold:true,color:ROSE,margin:0,isTextBox:true});
s.addText('«Las bolivianas merecen acceso a los mismos productos internacionales que el resto del mundo: originales, garantizados, y con alguien que sepa explicarte por qué uno funciona mejor que otro para tu tipo de piel».',{x:0.75,y:2.05,w:3.8,h:2.6,fontFace:BF,fontSize:15,italic:true,color:INK,margin:0,valign:'top',isTextBox:true});
s.addShape(pres.ShapeType.roundRect,{x:5.2,y:1.55,w:4.3,h:3.3,rectRadius:0.12,fill:{color:BLUSH},line:{color:BLUSH},objectName:'Tarjeta vision'});
s.addText('VISIÓN ACTUAL',{x:5.45,y:1.7,w:3.8,h:0.3,fontFace:BF,fontSize:13,bold:true,color:ROSE,margin:0,isTextBox:true});
s.addText('«La cadena de belleza especializada más grande de Bolivia» y «la primera cadena de tiendas para la mujer de Bolivia».',{x:5.45,y:2.05,w:3.8,h:2.6,fontFace:BF,fontSize:15,italic:true,color:INK,margin:0,valign:'top',isTextBox:true});
foot(s,3);

// 4 Qué falta
s=pres.addSlide(); title(s,'Paso 1: qué le falta a la misión y visión','Usamos las preguntas de análisis para encontrar las carencias');
const q=[['¿Quiénes somos?','Cadena especializada de belleza más grande del país','El lenguaje «girls» acota la identidad'],
['¿Qué hacemos?','Productos originales de 50+ marcas, con testers','No destaca el K-Beauty, su mayor diferencial'],
['¿Para quién?','Las bolivianas','No define segmentos ni menciona a los hombres'],
['¿En qué nos diferenciamos?','Originalidad, asesoría, tocar y probar, precios únicos','Está disperso en varios textos'],
['¿Por qué existimos?','Había un vacío entre la perfumería formal y lo informal','No está formalizado'],
['¿Adónde queremos llegar?','Crecer a más ciudades','No hay año ni meta medible']];
s.addTable([[hdr('Pregunta'),hdr('Lo que ya dice NUBA'),hdr('Lo que falta')]].concat(q.map((r,i)=>[cell(r[0],{bold:true,fill:{color:i%2?WHITE:BLUSH}}),cell(r[1],{fill:{color:i%2?WHITE:BLUSH}}),cell(r[2],{fill:{color:i%2?WHITE:BLUSH}})])),
 {x:0.5,y:1.5,w:9,colW:[2.4,3.3,3.3],rowH:0.52,border:{type:'none'},valign:'middle',margin:[0.04,0.1,0.04,0.1]});
foot(s,4);

// 5 y 6: de la actual a la mejorada
function mejora(num,t,sub,actual,mejor,cambios){
 s=pres.addSlide(); title(s,t,sub);
 s.addShape(pres.ShapeType.roundRect,{x:0.5,y:1.45,w:4.1,h:1.75,rectRadius:0.1,fill:{color:BLUSH},line:{color:BLUSH},objectName:'Actual '+num});
 s.addText('ACTUAL',{x:0.7,y:1.52,w:3.7,h:0.25,fontFace:BF,fontSize:12,bold:true,color:ROSE,margin:0,isTextBox:true});
 s.addText(actual,{x:0.7,y:1.8,w:3.7,h:1.35,fontFace:BF,fontSize:12,italic:true,color:INK,margin:0,valign:'top',isTextBox:true});
 s.addText('→',{x:4.62,y:2.0,w:0.76,h:0.5,fontFace:BF,fontSize:30,bold:true,color:BERRY,align:'center',margin:0,isTextBox:true});
 s.addShape(pres.ShapeType.roundRect,{x:5.4,y:1.45,w:4.1,h:1.75,rectRadius:0.1,fill:{color:BERRY},line:{color:BERRY},objectName:'Mejorada '+num});
 s.addText('MEJORADA',{x:5.6,y:1.52,w:3.7,h:0.25,fontFace:BF,fontSize:12,bold:true,color:'F3D9E0',margin:0,isTextBox:true});
 s.addText(mejor,{x:5.6,y:1.8,w:3.7,h:1.35,fontFace:BF,fontSize:12,color:WHITE,margin:0,valign:'top',isTextBox:true});
 s.addTable([[hdr('Lo que faltaba'),hdr('Lo que cambiamos')]].concat(cambios.map((r,i)=>[cell(r[0],{fill:{color:i%2?WHITE:BLUSH}}),cell(r[1],{bold:true,fill:{color:i%2?WHITE:BLUSH}})])),
  {x:0.5,y:3.4,w:9,colW:[4.2,4.8],rowH:0.36,border:{type:'none'},valign:'middle',margin:[0.03,0.1,0.03,0.1]});
 foot(s,num);}
mejora(5,'Paso 2: de la misión actual a la mejorada','Cada cambio responde a una carencia',
 '«Las bolivianas merecen acceso a los mismos productos internacionales que el resto del mundo: originales, garantizados, y con alguien que sepa explicarte por qué uno funciona mejor que otro para tu tipo de piel».',
 'Hacer que la belleza sea accesible, auténtica y cercana para las bolivianas y los bolivianos, con productos originales de más de 50 marcas, asesoría experta y una experiencia donde se puede tocar, probar y descubrir, con los mismos precios en todo el país.',
 [['No define segmentos ni menciona a los hombres','«las bolivianas y los bolivianos»'],['Diferenciación dispersa en varios textos','Originales, asesoría, tocar y probar, mismos precios'],['Razón de existir no formalizada','Belleza «accesible, auténtica y cercana»']]);
mejora(6,'Paso 3: de la visión actual a la mejorada','Valores: autenticidad · cercanía · inclusión · curaduría · innovación · responsabilidad',
 '«La cadena de belleza especializada más grande de Bolivia» y «la primera cadena de tiendas para la mujer de Bolivia».',
 'Ser en 2030 la cadena de belleza de referencia de Bolivia y la comunidad de skincare y maquillaje favorita de las nuevas generaciones: líder en K-Beauty, presente en las principales ciudades y en internet, con una marca propia reconocida.',
 [['No hay año ni meta medible','«en 2030» y metas que medimos en el Capítulo 4'],['No destaca el K-Beauty','«líder en K-Beauty»'],['Solo habla de crecer a más ciudades','Ciudades, internet y marca propia (Nu)']]);

// 7 Preguntas contestadas
s=pres.addSlide(); title(s,'Paso 4: contestamos las preguntas','Con la misión y visión mejoradas');
const r=[['¿Quiénes somos?','La cadena de belleza de referencia de Bolivia y su comunidad de skincare y maquillaje'],
['¿Qué hacemos?','Vendemos productos originales de 50+ marcas, con asesoría experta y K-Beauty'],
['¿Para quién?','Las bolivianas y los bolivianos, sobre todo las nuevas generaciones'],
['¿En qué nos diferenciamos?','Originalidad, asesoría, tocar y probar, mismos precios y marca propia'],
['¿Por qué existimos?','Para que la belleza sea accesible, auténtica y cercana, entre lo formal y lo informal'],
['¿Adónde queremos llegar?','A ser en 2030 la referencia del país, en las principales ciudades y en internet']];
s.addTable([[hdr('Pregunta'),hdr('Respuesta')]].concat(r.map((x,i)=>[cell(x[0],{bold:true,fill:{color:i%2?WHITE:BLUSH}}),cell(x[1],{fill:{color:i%2?WHITE:BLUSH}})])),
 {x:0.5,y:1.5,w:9,colW:[2.8,6.2],rowH:0.52,border:{type:'none'},valign:'middle',margin:[0.04,0.1,0.04,0.1]});
foot(s,7);

// 6 Matriz de poder
s=pres.addSlide(); title(s,'Grupos de interés: matriz de poder','Poder vs. interés de 1 a 10 (valoración del grupo)');
const quad=[
 ['PROMOTORES · mucho poder, mucho interés','Familia fundadora, Gerencia comercial, Marcas proveedoras, Clientas jóvenes, Competidores',0,0,BERRY],
 ['LATENTES · mucho poder, poco interés','Ventura Mall, AGEMED, Aduana e Impuestos, Banco Central, Influencers',1,0,ROSE],
 ['DEFENSORES · poco poder, mucho interés','Asesoras de tienda, Logística y courier',0,1,GREEN],
 ['INDIFERENTES · poco poder, poco interés','Mercado informal, Medios, Consumidores organizados',1,1,GREY]];
quad.forEach(([h,b,cx,cy,col])=>{const x=0.5+cx*4.55,y=1.5+cy*1.8;
 s.addShape(pres.ShapeType.roundRect,{x,y,w:4.45,h:1.7,rectRadius:0.1,fill:{color:BLUSH},line:{color:BLUSH},objectName:'Cuadrante '+h.split(' ')[0]});
 s.addText(h,{x:x+0.2,y:y+0.12,w:4.05,h:0.3,fontFace:BF,fontSize:12,bold:true,color:col,margin:0,isTextBox:true});
 s.addText(b,{x:x+0.2,y:y+0.5,w:4.05,h:1.1,fontFace:BF,fontSize:14,color:INK,margin:0,valign:'top',isTextBox:true});});
foot(s,6);

// 7 Cómo gestionarlos
s=pres.addSlide(); title(s,'Cómo gestionar a los actores clave');
const act=[['Familia fundadora','Apoya','9 / 9','Alinear la estrategia y las metas'],['Marcas proveedoras','Apoya','9 / 7','Cumplir metas de compra y cuidar exclusividades'],['AGEMED','Neutral','9 / 3','Tener los registros sanitarios al día'],['Aduana e Impuestos','Neutral','8 / 4','Cumplir trámites; apoyarse en un agente aduanero'],['Competidores','Se opone','6 / 8','Diferenciarse sin confrontar'],['Mercado informal','Se opone','5 / 3','Campaña «original garantizado» y denuncias']];
s.addTable([[hdr('Actor'),hdr('Postura'),hdr('Poder / interés'),hdr('Qué hacer')]].concat(act.map((r,i)=>r.map(t=>cell(t,{fill:{color:i%2?WHITE:BLUSH}})))),
 {x:0.5,y:1.2,w:9,colW:[2.3,1.2,1.5,4.0],rowH:0.5,border:{type:'none'},valign:'middle',margin:[0.04,0.1,0.04,0.1]});
foot(s,7);

// 8 PESTEL chart
s=pres.addSlide(); title(s,'PESTEL: perfil estratégico del entorno','Nota promedio por dimensión (1 amenaza · 5 oportunidad)');
s.addChart(pres.charts.BAR,[{name:'Nota promedio',labels:['Económica','Político-legal','Ecológica','Demográfica','Tecnológica','Sociocultural'],values:[1.9,2.3,3.6,4.0,4.0,4.1]}],
 {x:0.5,y:1.4,w:5.8,h:3.7,barDir:'bar',chartColors:[BERRY],showValue:true,dataLabelFormatCode:'0.0',dataLabelPosition:'outEnd',dataLabelFontFace:BF,dataLabelFontSize:12,dataLabelColor:INK,catAxisLabelFontFace:BF,catAxisLabelFontSize:12,catAxisLabelColor:INK,valAxisLabelColor:GREY,valAxisLabelFontFace:BF,valAxisMinVal:0,valAxisMaxVal:5,valGridLine:{color:'E6D6DB',size:0.5},catGridLine:{style:'none'},showLegend:false});
s.addText('3,0',{x:6.8,y:1.7,w:2.7,h:1.2,fontFace:HF,fontSize:72,bold:true,color:BERRY,margin:0,isTextBox:true});
s.addText('resultado total de 5: casi neutral',{x:6.8,y:2.9,w:2.7,h:0.6,fontFace:BF,fontSize:14,color:GREY,margin:0,isTextBox:true});
s.addText('Lo económico pesa más (30%) y está en rojo.',{x:6.8,y:3.7,w:2.7,h:0.8,fontFace:BF,fontSize:14,bold:true,color:RED,margin:0,isTextBox:true});
foot(s,8);

// 9 PESTEL factores
s=pres.addSlide(); title(s,'Factores clave del entorno');
const th={text:'',options:{}};
function col(x,head,color,items){
 s.addShape(pres.ShapeType.roundRect,{x,y:1.2,w:4.45,h:3.8,rectRadius:0.1,fill:{color:BLUSH},line:{color:BLUSH},objectName:head});
 s.addText(head,{x:x+0.2,y:1.3,w:4.05,h:0.35,fontFace:BF,fontSize:16,bold:true,color,margin:0,isTextBox:true});
 s.addText(items.map((t,i)=>({text:t,options:{bullet:true,breakLine:i<items.length-1}})),{x:x+0.2,y:1.75,w:4.05,h:3.1,fontFace:BF,fontSize:14,color:INK,margin:0,valign:'top',paraSpaceAfter:5,isTextBox:true});}
col(0.5,'Amenazas',RED,['Fin del dólar fijo: dólar casi 40% más caro (nota 1)','Inflación y poder de compra (2)','Bloqueos y conflictividad social (2)','Aranceles y trámites de aduana (2)','Consumo más selectivo (2)']);
col(5.05,'Oportunidades',GREEN,['Redes sociales y moda K-Beauty (5)','Población joven y urbana (4)','Santa Cruz, el departamento más poblado (4)','Pagos QR, e-commerce y WhatsApp (4)','Crecimiento del mercado y productos veganos (4)']);
foot(s,9);

// 10 Porter
s=pres.addSlide(); title(s,'Cinco fuerzas de Porter','Borrador del equipo a partir de los datos del trabajo; validar con la docente');
const forces=[['Rivalidad entre competidores','Media-alta','Otras cadenas de belleza; NUBA se diferencia por originalidad, asesoría y K-Beauty'],
['Nuevos entrantes','Media','Barreras: registro AGEMED, exclusividades de marca, aduana y dólar caro'],
['Productos sustitutos','Media-alta','Mercado informal y compras por internet, más baratos pero sin garantía'],
['Poder de proveedores','Alta','Marcas con poder 9/10, metas de compra y compras en dólares'],
['Poder de clientes','Media-alta','Consumo más selectivo, comparan precios y reponen menos']];
s.addTable([[hdr('Fuerza'),hdr('Intensidad'),hdr('Por qué')]].concat(forces.map((r,i)=>[cell(r[0],{bold:true,fill:{color:i%2?WHITE:BLUSH}}),cell(r[1],{bold:true,color:BERRY,fill:{color:i%2?WHITE:BLUSH}}),cell(r[2],{fill:{color:i%2?WHITE:BLUSH}})])),
 {x:0.5,y:1.5,w:9,colW:[2.4,1.4,5.2],rowH:0.6,border:{type:'none'},valign:'middle',margin:[0.04,0.1,0.04,0.1]});
foot(s,10);

// 11 Poder y rentabilidad
s=pres.addSlide(); title(s,'Poder y rentabilidad del sector');
[['Poder','Los proveedores (marcas) y las clientas informadas concentran el poder de negociación.'],['Rentabilidad','Moderada: la demanda joven crece, pero el dólar y la inflación aprietan los márgenes.'],['Qué hacer','Marca propia Nu, exclusividades y tienda online para recuperar margen.']].forEach(([h,b],i)=>{
 const x=0.5+i*3.05;
 s.addShape(pres.ShapeType.roundRect,{x,y:1.5,w:2.9,h:3,rectRadius:0.1,fill:{color:BLUSH},line:{color:BLUSH},objectName:'Tarjeta '+h});
 s.addText(h,{x:x+0.2,y:1.65,w:2.5,h:0.45,fontFace:HF,fontSize:22,bold:true,color:BERRY,margin:0,isTextBox:true});
 s.addText(b,{x:x+0.2,y:2.2,w:2.5,h:2.1,fontFace:BF,fontSize:14,color:INK,margin:0,valign:'top',isTextBox:true});});
foot(s,11);

// 12 Conclusión
s=pres.addSlide(); s.background={color:BERRY};
s.addText('Conclusión',{x:0.5,y:0.4,w:9,h:0.7,fontFace:HF,fontSize:36,bold:true,color:WHITE,margin:0,isTextBox:true});
s.addText([{text:'NUBA tiene aliados fuertes: marcas, clientas jóvenes y su propia dirección.',options:{bullet:true,breakLine:true}},
{text:'Del entorno le ayuda la demanda: jóvenes, redes sociales y pagos digitales.',options:{bullet:true,breakLine:true}},
{text:'Le perjudican los costos: dólar más caro, inflación, aduana y bloqueos.',options:{bullet:true,breakLine:true}},
{text:'Perfil del entorno: 3,0 de 5.',options:{bullet:true}}],{x:0.5,y:1.5,w:9,h:3,fontFace:BF,fontSize:20,color:WHITE,margin:0,valign:'top',paraSpaceAfter:12,isTextBox:true});

// 13 O-A
s=pres.addSlide(); title(s,'Oportunidades y amenazas');
col(0.5,'Oportunidades',GREEN,['K-Beauty y redes sociales','Mercado joven y urbano en Santa Cruz','Pagos QR y comercio electrónico','Crecimiento de cosméticos (≈6% anual)']);
col(5.05,'Amenazas',RED,['Dólar casi 40% más caro','Inflación y menor poder de compra','Bloqueos y abastecimiento','Aduana y regulación sanitaria']);
foot(s,13);

pres.writeFile({fileName:'NUBA_Presentacion.pptx'}).then(()=>console.log('ok'));
