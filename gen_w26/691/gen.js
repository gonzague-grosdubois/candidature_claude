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
"LettreMotivation_CreditAgricoleSA_AnalysteEnergieFinance_Paris.docx":["Madame, Monsieur,",
"Je souhaite vous proposer ma candidature au stage d'Analyste Energie et Finance au sein de la Direction Solutions Durables et Décarbonation de Crédit Agricole S.A. à Paris. Préparer des pitchs sectoriels pour les clients et prospects de la filière énergie, assurer la veille sur les énergies renouvelables, la mobilité électrique, les smartgrids et l'autoconsommation, préparer les rendez vous du Head of Energy, suivre les prix et volumes de l'énergie, du spot aux PPA, et accompagner la croissance des entreprises par une ingénierie financière mêlant dette et equity est exactement le travail qui relie mon intérêt pour les marchés et pour la transition énergétique.",
P2FR,
"Mon Master 1 Finance à l'ESSCA, où j'étudie la gestion quantitative des risques, la gestion de portefeuille et les produits dérivés, ainsi que ma certification Bloomberg Market Concepts, m'ont donné des bases solides en analyse financière et en marchés. "+AI_FR+" Ces projets m'ont appris à suivre des prix et des volumes, à structurer des données et à tirer une conclusion claire d'un jeu de chiffres, ce que j'apporterais à vos outils de monitoring et à vos pitchs sectoriels.",
"Mon parcours commercial m'a appris à préparer un rendez vous et à convaincre. "+COM_FR,
DISPO_FR,
"Je serais ravi de rejoindre l'équipe Energie de Crédit Agricole S.A. et me tiens à votre disposition pour un entretien.",...CLOSEFR]};
(async()=>{for(const [f,paras] of Object.entries(letters)){
const text=paras.join("\n");
const h=text.replace(/GROS-DUBOIS/g,"").match(/[^\s]*-[^\s]*/g),d=text.match(/[—–]/g),b=text.match(/avenant|amendment|écourter|shorten|flexib|conflict/gi);
console.log(f,'Hyphens:',h,'Dashes:',d?d.length:0,'Banned:',b);
if(h||d||b) process.exit(1);
const doc=new Document({sections:[{properties:{page:{size:{width:11906,height:16838},margin:{top:1134,bottom:1134,left:1134,right:1134}}},
children:paras.map(t=>new Paragraph({alignment:AlignmentType.JUSTIFIED,spacing:{after:200},children:[new TextRun({text:t,font:"Arial",size:22})]}))}]});
fs.writeFileSync('/tmp/lettres_txt/'+f.replace('.docx','')+'.txt',text+'\n');fs.writeFileSync('/home/claude/cc/lettres/'+f,await Packer.toBuffer(doc));console.log('ok',f);}})();
