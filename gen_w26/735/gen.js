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
const DISPO_EN="I am available for the summer 2027 programme, from June 2027.";
const letters={"CoverLetter_AdamsStreet_Primary_MenloPark.docx":["Dear Hiring Team,",
"I am writing to apply for the Primary Intern position in the Summer Analyst Program 2027 at Adams Street Partners in Menlo Park. Contributing to firm and client projects, underwriting and performing due diligence on investments, and conducting sector and thematic research within a US Venture team with exposure to buyouts is exactly the way I want to learn private markets investing.",
P2EN,
"My Master's 1 in Finance at ESSCA, where I study portfolio management, quantitative risk management and IFRS, and my Bloomberg Market Concepts certification have given me solid foundations in investment analysis and Excel modelling. "+AI_EN+" These projects taught me to test a thesis against data and to use AI tools to research faster, which I would bring to your underwriting and sector research.",
"My commercial experience taught me to evaluate an opportunity, negotiate and close. "+COM_EN,
DISPO_EN,
"I would be delighted to join the Primary team of Adams Street Partners this summer and would welcome the opportunity of an interview.",...CLOSEEN]};
(async()=>{for(const [f,paras] of Object.entries(letters)){
const text=paras.join("\n");
const h=text.replace(/GROS-DUBOIS/g,"").match(/[^\s]*-[^\s]*/g),d=text.match(/[—–]/g),b=text.match(/avenant|amendment|écourter|shorten|flexib|conflict/gi);
console.log(f,'Hyphens:',h,'Dashes:',d?d.length:0,'Banned:',b);
if(h||d||b) process.exit(1);
const doc=new Document({sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:1134,bottom:1134,left:1134,right:1134}}},
children:paras.map(t=>new Paragraph({alignment:AlignmentType.JUSTIFIED,spacing:{after:200},children:[new TextRun({text:t,font:"Arial",size:22})]}))}]});
fs.writeFileSync('/tmp/lettres_txt/'+f.replace('.docx','')+'.txt',text+'\n');fs.writeFileSync('/home/claude/cc/lettres/'+f,await Packer.toBuffer(doc));console.log('ok',f);}})();
