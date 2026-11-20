const fs=require('fs');
const {Document,Packer,Paragraph,TextRun,AlignmentType}=require('docx');
const P2FR="Je réalise actuellement un stage de six mois en tant qu'Assistant Patrimonial chez L&A Finance à Paris, cabinet de conseil en gestion de patrimoine du Groupe Rayne, où j'analyse et structure des transactions financières sur différents produits d'investissement, conduis une veille de marché quotidienne et contribue à la rédaction de supports techniques à destination des clients.";
const P2EN="I am currently completing a six month internship as Wealth Management Intern at L&A Finance in Paris, an independent wealth management advisory firm within the Groupe Rayne, where I analyse and structure financial transactions across various investment products, conduct daily market monitoring and contribute to the preparation of technical client materials.";
const CLOSEFR=["Dans l'attente de votre retour, veuillez agréer, Madame, Monsieur, l'expression de mes salutations distinguées.","","Gonzague GROS-DUBOIS"];
const CLOSEEN=["Yours sincerely,","","Gonzague GROS-DUBOIS"];
const AI_EN="Using AI tools, I designed a quantitative equity strategy that achieved a Sharpe ratio of 0.91 against the SPY over a ten year backtest, and I built a trading system connected to eight exchanges with real time P&L monitoring, which I shut down once the data showed no net edge.";
const AI_FR="En m'appuyant sur des outils d'IA, j'ai conçu une stratégie algorithmique atteignant un ratio de Sharpe de 0,91 par rapport au SPY sur dix ans de backtest, et un système de trading connecté à huit plateformes avec un suivi du P&L en temps réel, que j'ai arrêté lorsque les données ont montré l'absence d'avantage net.";
const COM_EN="At Mecanicus, I completed more than ten over the counter transactions between €50k and €500k and helped manage €3M of inventory, negotiating and documenting each deal. As an independent reseller since 2018, I generated over €10k of revenue with a net margin of around 50%, prospecting and closing on my own. I am fluent in English (TOEIC 975/990) and French, and I also speak Spanish.";
const COM_FR="Chez Mecanicus, j'ai réalisé plus de dix transactions de gré à gré entre 50 000 et 500 000 euros et participé à la gestion d'un inventaire de 3 millions d'euros, en négociant et en documentant chaque opération. Revendeur indépendant depuis 2018, j'ai généré plus de 10 000 euros de chiffre d'affaires avec une marge nette d'environ 50 %, en prospectant et en concluant seul mes ventes. Je parle couramment anglais (TOEIC 975/990) et français, ainsi que l'espagnol.";
const DISPO_EN="I am available from January 2027 for six months.";
const letters={
"CoverLetter_OCBC_EquityResearch_Singapore.docx":["Dear Hiring Team,",
"I would like to apply for the Global Markets Research internship in Equity Research at OCBC in Singapore. Researching listed companies and their performance, positioning and growth, analysing sectors, building and updating valuation models, following announcements, results and reports, writing research notes and market commentary, and supporting Global Markets, Wealth, Private Banking and Securities is exactly the way I want to learn the equity research craft.",
P2EN,
"My Master's in Finance at ESSCA, where I study portfolio management, derivative pricing and IFRS, together with my Bloomberg Market Concepts certification, gave me a solid base in valuation and fundamental analysis. "+AI_EN+" These projects taught me to combine fundamental and technical reading of the markets, to test an idea against data and to write down a clear conclusion, which I would bring to your valuation models and research notes.",
"My commercial background taught me to form a view on a price, to defend it and to close. "+COM_EN,
DISPO_EN,
"I would be delighted to join the Equity Research team at OCBC and remain at your disposal for an interview.",...CLOSEEN]};
(async()=>{for(const [f,paras] of Object.entries(letters)){
const text=paras.join("\n");
const h=text.replace(/GROS-DUBOIS/g,"").match(/[^\s]*-[^\s]*/g),d=text.match(/[—–]/g),b=text.match(/avenant|amendment|écourter|shorten|flexib|conflict/gi);
console.log(f,'Hyphens:',h,'Dashes:',d?d.length:0,'Banned:',b);
if(h||d||b) process.exit(1);
const doc=new Document({sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:1134,bottom:1134,left:1134,right:1134}}},
children:paras.map(t=>new Paragraph({alignment:AlignmentType.JUSTIFIED,spacing:{after:200},children:[new TextRun({text:t,font:"Arial",size:22})]}))}]});
fs.writeFileSync('/tmp/lettres_txt/'+f.replace('.docx','')+'.txt',text+'\n');fs.writeFileSync('/home/claude/cc/lettres/'+f,await Packer.toBuffer(doc));console.log('ok',f);}})();
