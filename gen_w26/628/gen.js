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
const DISPO_FR="Je suis disponible à partir de janvier 2027 pour une durée de six mois.";
const letters={
"LettreMotivation_SocieteGenerale_AnalysteCoverageIBFrance_LaDefense.docx":["Madame, Monsieur,",
"Je souhaite vous proposer ma candidature au stage d'Analyste Coverage et Investment Banking France, division Automotive, Mobilities, Defence and Aerospace, au sein de Société Générale CIB à La Défense. Suivre la relation globale avec de grands clients français de l'aéronautique, de l'automobile et de la défense, soutenir les actions commerciales, aider à rédiger les présentations destinées à leur top management avec les lignes métiers (M&A, ECM, DCM, Cash Management, marchés), simuler la rentabilité et le capital mobilisé des nouvelles opérations de banque commerciale, suivre les RWA du portefeuille et animer la revue de presse sectorielle est exactement la façon dont je veux apprendre le métier de banquier coverage.",
P2FR,
"Mon Master 1 Finance à l'ESSCA, où j'étudie la gestion quantitative des risques, la modélisation des produits de taux et la gestion de portefeuille, ainsi que ma certification Bloomberg Market Concepts, m'ont donné des bases solides pour lire la rentabilité d'un client et le capital qu'il consomme. "+AI_FR+" Ces projets m'ont appris à structurer des données, à synthétiser des résultats et à tester une hypothèse avant de la présenter, ce que j'apporterais aux simulations de rentabilité, au suivi des RWA et à la préparation des présentations clients.",
"Mon parcours commercial m'a appris à écouter un client, à négocier et à conclure. "+COM_FR,
DISPO_FR,
"Je serais ravi de rejoindre l'équipe Coverage Automotive, Mobilities, Defence and Aerospace de Société Générale CIB et me tiens à votre disposition pour un entretien.",...CLOSEFR]};
(async()=>{for(const [f,paras] of Object.entries(letters)){
const text=paras.join("\n");
const h=text.replace(/GROS-DUBOIS/g,"").match(/[^\s]*-[^\s]*/g),d=text.match(/[—–]/g),b=text.match(/avenant|amendment|écourter|shorten|flexib|conflict/gi);
console.log(f,'Hyphens:',h,'Dashes:',d?d.length:0,'Banned:',b);
if(h||d||b) process.exit(1);
const doc=new Document({sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:1134,bottom:1134,left:1134,right:1134}}},
children:paras.map(t=>new Paragraph({alignment:AlignmentType.JUSTIFIED,spacing:{after:200},children:[new TextRun({text:t,font:"Arial",size:22})]}))}]});
fs.writeFileSync('/tmp/lettres_txt/'+f.replace('.docx','')+'.txt',text+'\n');fs.writeFileSync('/home/claude/cc/lettres/'+f,await Packer.toBuffer(doc));console.log('ok',f);}})();
