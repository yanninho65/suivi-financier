# Suivi financier — Documentation de référence

**Version documentée : `v 2026.09.12.23.19`** — à comparer à `APP_VERSION` en tête du `<script>` du fichier `index.html` (format `v AAAA.MM.JJ.HH.MM`, affichée en petit sous le titre "Suivi financier" sans le préfixe `v `). Mise à jour uniquement sur demande explicite : si les versions diffèrent, vérifier en priorité §2, §4, §5, §6.1, §6.2, §6.3, §6.4, §6.6, §6.7, §7.2 et §8 (zones les plus souvent modifiées).

Fichier livré : **`index.html`** (page unique, style One UI, vanilla JS). Un seul CDN au runtime : SheetJS (`xlsx`, import/export Excel). Aucun backend, aucune donnée envoyée nulle part.

Sert de visualisation/filtrage mobile d'un suivi Excel existant, saisi manuellement. Le classeur Excel de l'utilisateur reste la source de vérité ; import incrémental avec déduplication.

---

## 1. Modèle de données

### 1.1 Colonnes d'import (Excel/CSV, mapping par nom normalisé, ordre libre)
`Date · Montant · Type · Catégorie · Personne · Compte · Compte cible · Groupe · Remarque`

- **Date** (`U.parseDateValue`) : vraie cellule Excel, numéro de série Excel, texte `JJ/MM/AAAA`, ou texte `JJ/MM/AA` (année 2 chiffres, pivot 68 : `00`-`68` → 2000-2068, `69`-`99` → 1969-1999).
- **Type** (`U.normType`) : accepte l'interne anglais (`Income`/`Expense`/`Transfer`) et le français, insensible casse/accents — `revenu`, tout ce qui commence par `exp`/`dépense`, `virement`, ou tout ce qui commence par `trans`/`transfert`.
- **Montant** : nombre ou texte avec virgule décimale, signe porté directement par la valeur (positif = revenu, négatif = dépense) — jamais déduit de la colonne Type.

### 1.2 Objet transaction interne
```
{ id, date, dateISO, montant (signé), type, categorie, personne, compte, compteCible, groupe, remarque,
  groupeBase,   // "groupe" sans le suffixe "(...)" final, ex. "Cliff Haven (23-01-08)" -> "Cliff Haven"
  transform }   // { mode: "same"|"month"|"yearsplit", year, month }
```
`groupeBase` n'est qu'un helper d'affichage/de repli — voir §2 : le libellé complet (avec suffixe) est la clé de référence partout, jamais le nom de base seul.

### 1.3 Date transformée (`computeTransform()`)
Dérivée de la `Remarque` :
- suffixe `AA-MM` → transaction affectée à ce mois précis (`mode:"month"`).
- suffixe année à 4 chiffres → montant réparti également sur les 12 mois de l'année (`mode:"yearsplit"`). Dans un contexte de **période** (Budget, Évolution en mode "Mois transformé"), seule la part correspondant aux mois de la période est montrée, jamais le montant total d'origine.
- sinon → identique à la date indiquée (`mode:"same"`).

Ce triplet `{mode, year, month}` est calculé une fois par transaction (au chargement/import) et réutilisé tel quel par le moteur Budget (`computeGroupBudget`) et le moteur Évolution (`buildEvolutionIndex`) — même logique de redistribution mensuelle dans les deux onglets.

---

## 2. Groupes de groupe / groupes de personnes

Un même lieu/séjour donne souvent plusieurs groupes datés (`Cliff Haven (23-01-08)`, `Cliff Haven (24-02)`) qu'on veut rattacher à une période de vie plus large (`Ghana`). Règles éditables à la main, jamais déduites.

### 2.1 Principe central : le libellé complet prime toujours, jamais le nom de base

Deux occurrences distinctes du même nom de base (ex. "France (26-06_07)" et "France (25-07_08)") sont des groupes différents et ne doivent jamais se voir attribuer la même classification par accident :

- **`findGroupAliasRule(base, full)`** (règles `groupe → groupe de groupe`) : cherche d'abord une règle exacte sur `full` (libellé littéral complet, ex. `"France (26-06-07)"`) ; ce n'est que si aucune règle `full` n'existe qu'un repli sur une règle "base générique" (`groupe = nom de base, pas de full`) s'applique — et cette règle générique n'est alors utilisée que pour les occurrences non rattachées individuellement, jamais pour outrepasser une règle `full` existante. Pas de fenêtre de dates sur les règles : la seule façon de distinguer deux occurrences est le libellé complet lui-même.
- **`getGroupKind(fullName)` / `getGroupVoyageMonth(fullName)`** (classification Vie quotidienne/Voyage et mois d'affectation, §2.4) : indexées sur le libellé complet.
- **`getGroupDateRange(key)`** (dates de séjour → jours/moyennes, §2.4) : indexée sur le libellé complet (`editingFull || editingBase` côté UI).

En clair : un nom de base seul n'est jamais utilisé pour appliquer un réglage à une occurrence précise — il ne sert que de mécanisme volontaire et explicite de "règle générique" que l'utilisateur choisit de créer via la case "Appliquer à toutes les périodes de « X »" de l'écran d'édition d'un groupe (§2.4), pour classer en bloc toute occurrence non encore rattachée individuellement.

### 2.1 bis Tri chronologique des groupes de groupe

Partout où les groupes apparaissent groupés par groupe de groupe et où un ordre chronologique a du sens — picker Groupe de la fiche transaction (§6.2), vue Groupe de l'onglet Évolution (§6.4.1), filtre Groupe de l'onglet Transactions (§4), picker Groupe d'un ajustement manuel du Budget (§6.3.6) — le classement suit `AGG.groupeHierarchyChrono(transactions)`, variante triée de `groupeHierarchyFull` : les groupes de groupe sont classés **du plus récent au plus ancien**, en se basant sur la **date de début du groupe enfant le plus ancien** de chacun ; à l'intérieur d'un groupe de groupe, ses groupes enfants sont classés **du plus ancien au plus récent** (leur propre date de début). Les groupes non rattachés à un groupe de groupe restent à part (`flat`, triés alphabétiquement) — jamais fondus dans un groupe de groupe, jamais mélangés au tri chronologique. Le résultat porte un indicateur `preSorted:true` que les vues consommatrices respectent pour ne pas re-trier alphabétiquement par-dessus (`renderPickerList`/`nestedFilterHtml`).

Le picker de groupe **principal/additionnel** du Budget (`groupPickerHierarchy`, §6.3.2) n'est pas concerné — il utilise `groupeHierarchyFull`, trié alphabétiquement. Le picker de groupe d'un **ajustement manuel** (`adjustmentGroupHierarchy`, §6.3.6) réutilise `groupeHierarchyChrono` : les deux pickers de groupe du Budget suivent donc des tris différents selon leur rôle (principal/additionnel = alphabétique ; ajustement = chronologique, aligné sur la fiche transaction).

**Date de "départ" d'un groupe** (`AGG.getGroupStartDate(fullLabel, txnMinDatesMap)`), ordre de priorité strict :
1. La **date de début** configurée manuellement pour ce groupe (`getGroupDateRange`, §2.5) — toujours prioritaire quand elle existe.
2. À défaut, la **date entre parenthèses** dans le libellé, interprétée par `AGG.parseGroupSuffixDate` (via `AGG.groupLabelSuffix`) : `AA-MM-JJ` → jour précis (ex. `23-01-08` → 8 janvier 2023) ; `AA-MM_MM` → plage de mois sur une même année, départ = 1ᵉʳ du premier mois (ex. `24-01_06` → 1ᵉʳ janvier 2024) ; `AA-MM` → mois, départ = 1ᵉʳ du mois (ex. `22-07` → 1ᵉʳ juillet 2022) ; `AAAA` → année seule, départ = 1ᵉʳ janvier.
3. En dernier recours, la **date de la transaction la plus ancienne** de ce groupe précis (`AGG.groupStartDatesMap(transactions)`, calculée une fois pour tous les groupes).

### 2.2 Règles (`groupAliasRules` / `personAliasRules`)
```
{ groupe, superGroupe, full? }   // groupes : "full" = libellé complet si scopé à une occurrence précise
{ personne, superPersonne }      // personnes : jamais d'occurrences datées, donc jamais de champ "full"
```
`getGroupeDisplay(t)` = super-groupe si une règle correspond (via `findGroupAliasRule`, §2.1), sinon `t.groupe` brut (libellé complet, jamais tronqué).

### 2.3 Où ça s'applique
- **Filtres** Groupe/Personne : liste imbriquée, groupes de groupe repliés par défaut, cochable au niveau du cluster ou d'un élément précis (`Filters.matches()` teste la valeur résolue **et** la valeur brute).
- **Transactions**, regroupement "Compte" : utilise le **compte réel** (`t.compte`, nom exact) — le type de compte (§3) reste un mode de regroupement séparé.
- **Onglet Transactions (liste)** : le libellé affiché reste toujours le **groupe littéral** (`t.groupe`, avec date), jamais le groupe de groupe résolu — ligne 3 de chaque transaction, cf. §6.2.
- **Onglet Évolution** (§6.4), vue Groupe : sections = groupe de groupe résolu, items = libellé complet du groupe.

### 2.4 Écran "Gérer les groupes / Gérer les personnes"
Deux instances de `createAliasManager(cfg)` (`window.__SF_GROUPMGR` avec `hasDates:true, showKind:true`, `window.__SF_PERSONMGR` avec `hasDates:false`).

**Vue hiérarchique — points clés :**
- Chaque **occurrence datée distincte** apparaît comme sa propre ligne, avec son **libellé complet** (date incluse) — les groupes ne sont jamais fusionnés/masqués sous le seul nom de base. Idem pour la vue "Non classés".
- Chaque ligne affiche une icône (voyage/vie quotidienne) via `getGroupKind(b.base)` où `b.base` **est** le libellé complet.
- Position de scroll et sections `<details>` dépliées sont conservées entre deux re-rendus.

**Édition d'un groupe (tap sur une ligne) — édition par occurrence :**
Le tap ouvre un écran scopé à **cette occurrence précise** (pas à toutes les dates du même nom de base) :
- Le champ "groupe de groupe" est pré-rempli avec l'affectation qui s'applique réellement à cette occurrence (résolue via `findGroupAliasRule(editingBase, editingFull)`).
- **Enregistrer** crée par défaut une règle `full` scopée à cette occurrence précise — sans toucher aux règles des autres occurrences du même nom de base.
- Case à cocher optionnelle **"Appliquer à toutes les périodes de « X »"** : crée/remplace la règle générique (sans `full`) du nom de base — seul moyen de créer une règle générique dans l'interface (§2.1).
- Si le groupe n'a pas de suffixe de date (occurrence unique), pas de case à cocher : règle générique directe.
- **Type de groupe** (Vie quotidienne / Voyage) + mois voyage : `key = editingFull || editingBase`, indexé sur le libellé complet. Le mois voyage (`Store.settings.groupVoyageMonth`, format interne `AAAA-MM`) prime sur le **Mois calculé** partout dans l'app (regroupement Transactions §6.2, filtre §4 — voir `AGG.getCalcMonthYM(t)` ci-dessous). Export/import Excel (§8/§7) : la colonne "Mois voyage" écrit/lit une **vraie date Excel au 1ᵉʳ du mois** (`voyageMonthToDateObj`), jamais le texte `AAAA-MM` brut.

#### Mois calculé d'une transaction — `AGG.getCalcMonthYM(t)`
Fonction unique, seule source de vérité pour "à quel mois calculé appartient cette transaction", utilisée par le regroupement Transactions "Mois calculé" (§6.2) et le filtre "Mois calculé" (§4) — les deux doivent rester alignés sur cette même fonction. Ordre de priorité :
1. Si le groupe de la transaction est de type **"voyage"** (`getGroupKind`) **et** qu'un **mois voyage est renseigné** (`getGroupVoyageMonth`, non vide) → ce mois-là, pour toutes les transactions du groupe, quelle que soit leur date réelle ou leur Remarque.
2. Sinon, si `t.transform.mode==="month"` (suffixe `AA-MM` en Remarque, §1.3) → le mois transformé.
3. Sinon → le mois réel de la transaction (`t.date`).

Le mode `"yearsplit"` (répartition sur 12 mois, §1.3) n'est **jamais** éclaté ici — une transaction ne peut appartenir qu'à un seul groupe/mois à la fois dans une liste ou un filtre ; en `"yearsplit"` (ou `"same"`), on retombe sur le mois réel (règle 3).
- **Dates de début/fin** (§2.5) : disponibles pour tous les groupes, quel que soit leur type, même clé (`editingFull || editingBase`).

**Sélection multiple** (bouton "Sélectionner", **Personnes uniquement** — `!cfg.hasDates`, absent côté Groupes) : coche par nom de base réel, puis "Déplacer vers…" applique une règle générique en masse.

**Nettoyage des règles obsolètes** (`obsoleteRules()`) : un bouton "N règle(s) obsolète(s) · Nettoyer" apparaît en haut de l'écran uniquement s'il existe des règles dont le groupe/la personne visé·e n'apparaît plus dans **aucune** transaction actuelle — comparaison insensible casse/accents (`U.normalize`) sur le nom de base, ou sur le libellé complet exact pour une règle `full`. Cliquer ouvre une confirmation qui liste chaque règle concernée (`base/full → super-groupe`, jusqu'à 10 affichées, "+N de plus" au-delà) avant suppression définitive. Fonction générique, partagée par Groupes et Personnes.

> **Piège pour toute évolution de `createAliasManager`** : les deux instances (Groupes/Personnes) peuvent coexister dans le DOM avec des éléments internes de mêmes id (`am-close`, `am-search`, etc.). Tout nouveau `document.getElementById("am-…")`/`document.querySelectorAll("[data-…]")` ajouté dans une fonction de rendu doit être scopé au `sheet` local (`sheet.querySelector(...)`), jamais global à `document` — sinon le mauvais gestionnaire peut être câblé silencieusement sur l'autre écran.

### 2.5 Dates de début/fin d'un groupe → jours & moyennes/jour
`Store.settings.groupDateRange` : `{ [libellé complet]: {dateFrom, dateTo} }`. Saisi dans l'écran d'édition d'un groupe (tous types confondus).

Dès que les deux dates sont renseignées, calcul et affichage en direct :
- **Nombre de jours** = `(dateTo - dateFrom) + 1`.
- **Revenus/jour, Dépenses/jour, Total/jour** = totaux des transactions de ce groupe précis (`t.groupe === clé`) ÷ nombre de jours.
- Si une date manque, ou si `dateTo < dateFrom` (jours ≤ 0) : aucun calcul — message "Renseignez les deux dates…" à la place. `getGroupDayStats(key)` (exposé sur `AGG`) retourne `null` dans ces cas.

**Date de fin dynamique ("toujours aujourd'hui")** : case à cocher sous les deux champs de date, pour un groupe/séjour encore en cours. Cochée, le champ "Date de fin" se grise et affiche la date du jour (lecture seule) ; c'est la valeur spéciale `AGG.GROUP_DATE_TODAY_SENTINEL` (chaîne `"today"`, jamais une vraie date ISO) qui est stockée dans `dateTo`. Partout où un calcul a besoin d'une vraie date, `AGG.resolveGroupDateToValue(dateTo)` résout cette sentinelle en date du jour **au moment du calcul** ; ne s'applique qu'à `dateTo`, jamais à `dateFrom`. Export/import Excel (§7/§8) : la sentinelle s'écrit `"Aujourd'hui"` dans la cellule "Date fin (jours)" et se reconnaît à la réimport (comparaison insensible casse/accents sur `"aujourd'hui"`/`"today"`) — toute autre valeur est interprétée comme une vraie date via `U.parseDateValue`.

Exporté/importé avec les autres paramètres (§7). Il n'existe pas de fenêtre de dates sur les règles de groupe de groupe (`groupAliasRules`/`personAliasRules`) — seul `groupDateRange` porte une notion de dates, pour un usage différent : les jours de séjour réels, pas une fenêtre de validité de règle.

---

## 3. Types de compte et acronymes
`Store.settings.accountTypes` = `{ "Nom du compte": "Courant"|"Épargne"|texte libre }`, saisie manuelle uniquement. Compte sans type → "Courant" par défaut. Utilisé pour regrouper les comptes par type dans l'onglet Comptes (§6.1) et dans le regroupement "Compte" de Transactions (§6.2).

`Store.settings.accountAcronyms` = `{ "Nom du compte": "ACRO" }` (4 caractères max, majuscules). Réglable dans l'écran "Types de compte" (menu ☰), sous le type — `AGG.getAccountAcronym(compte)` retourne l'acronyme saisi, ou à défaut le dérive des initiales du nom (`deriveAccountAcronym`, jusqu'à 3 lettres). Affiché en **bandeau vertical** sur chaque transaction dans l'onglet Transactions (§6.2) — seule indication visuelle du compte dans la liste.

---

## 4. Filtres

Champs, dans l'ordre du menu (`MENU_ITEMS`) : **Période, Mois indiqué, Mois calculé, Type, Compte, Groupe, Catégorie, Personne**, plus une recherche texte libre. Cumulables entre champs (AND), multi-sélection en OR au sein d'un champ. Ouverture en deux étapes (menu de catégories de filtre → liste dédiée, `sheetView`). Catégorie triée Revenus A→Z puis Dépenses A→Z (`AGG.distinctValuesCategorieSorted`). Groupe : groupes de groupe triés chronologiquement (§2.1 bis), pas alphabétiquement.

**Filtres actifs** : visibles dans la sheet de filtres elle-même, sous le menu principal (`activeFiltersSectionHtml()`), regroupés par type de filtre (`CHIP_KIND_LABEL`/`CHIP_KIND_ORDER`), chaque puce individuellement supprimable (`F.removeChip`, re-render local sans fermer la sheet). Sur la page Transactions : badge numérique sur l'icône filtre du bouton flottant (§6.2, `updateFilterBadge`).

**Tout sélectionner / Aucun** : chaque section à choix multiples (Mois indiqué, Mois calculé, Type, Compte, Groupe, Catégorie, Personne — pas Période) affiche deux boutons en haut (`quickSelectRowHtml`/`wireQuickSelect`/`selectAllForKey`/`selectNoneForKey`).

**Mois indiqué / Mois calculé** (`Filters.moisIndiques`/`moisCalcules`, valeurs `"AAAA-MM"`) : listent respectivement les mois de la date réelle et du **mois calculé** (§2.4, `AGG.getCalcMonthYM`). Affichage **groupé par année** (`moisGroupedHtml`/`wireMoisGrouped`) : cliquer sur l'en-tête d'une année sélectionne/désélectionne en bloc tous ses mois — c'est le mécanisme pour filtrer sur une année entière, il n'existe pas de champ "année" séparé.

**Compte** : groupé par type de compte (§3, `comptesGroupedHtml`/`wireComptesGrouped`) — cliquer sur l'en-tête d'un type sélectionne/désélectionne tous les comptes de ce type (bascule directe de chaque membre, pas de valeur "type" unique dans `Filters.matches`).

**Groupe / Personne** : liste imbriquée (`nestedFilterHtml`/`wireNestedFilter`), groupes de groupe/personnes repliés par défaut, cochables au niveau du cluster ou d'un élément précis. **Personne** uniquement : les personnes non rattachées apparaissent sous un en-tête dédié **"Sans type"** (`NESTED_DIM_CONFIG.personnes.noGroupLabel`), cliquable pour tout sélectionner en bloc (bascule directe de chaque membre).

---

## 5. Réglage global des décimales
`Store.settings.decimals` (0 ou 2) : réglage **global**, partagé entre Transactions et Comptes — un bouton pilule "0,00" (`U.decimalsToggleBtnHtml`/`U.toggleMoneyDecimals`), présent dans l'onglet **Comptes** (§6.1), l'active/désactive pour les deux onglets. Inactif par défaut = 0 décimale.

**Exceptions** : l'onglet Budget affiche toujours 0 décimale, indépendamment de ce réglage. L'onglet Évolution aussi, via son propre formateur local (`fmtVal`) — indépendant de `Store.settings.decimals`.

**À ne pas confondre avec** la pilule "0€" (§6.1) : masque uniquement l'affichage des comptes à solde nul dans l'onglet Comptes (`Store.settings.hideZeroAccounts`), sans lien avec le nombre de décimales.

---

## 6. Onglets
Cinq onglets dans la barre de navigation : **Comptes** (par défaut au démarrage), **Transactions**, **Budget**, **Évolution**, **Récurrent** — plus deux onglets secondaires derrière un bouton "⋯" en fin de barre (§11, §12) : **Revenus** (§6.7) puis **Santé** (§6.6), dans cet ordre dans `MORE_TABS`.

### 6.1 Comptes
Solde total (tous comptes), puis les comptes regroupés par type (Courant / Épargne / types libres), chaque groupe affichant son solde. Chaque compte n'affiche que son nom et son solde.

Une seule ligne de boutons sous la carte de solde (`display:flex;gap:8px;flex-wrap:nowrap`) :
- **Copier pour Excel** (`buildAccountsClipboardText` + `copyTextToClipboard`) : copie dans le presse-papiers un texte tabulé (TSV) `Compte / Solde` — une ligne par compte (2 décimales, virgule, sans séparateur de milliers), plus une ligne "Total". Collé dans Excel, forme directement deux colonnes. Repli sur `document.execCommand("copy")` si l'API Clipboard moderne échoue (ex. ouverture en `file://`). Toujours calculée sur l'ensemble des comptes, indépendamment du filtre "masquer les comptes à 0 €".
- **Réorganiser** (visible dès qu'il y a plus d'un groupe de type) : pop-up avec flèches haut/bas par groupe (`openAccountTypeReorderSheet`). Ordre mémorisé dans `Store.settings.accountTypeOrder`, persisté entre sessions, inclus dans la sauvegarde `.json`/Gist (pas dans l'Excel, §7.2).
- Bouton décimales (§5).
- **Masquer les comptes à 0 €** (pilule "0€", `U.hideZeroAccountsBtnHtml`/`U.toggleHideZeroAccounts`) : filtre les comptes dont le solde arrondit à 0,00 € au centime près (`U.isZeroBalance(bal)`). Réglage local à l'appareil (`Store.settings.hideZeroAccounts`, jamais synchronisé, §7.2). Un groupe de type sans compte visible une fois le filtre appliqué n'est pas affiché.

Taper sur un compte filtre les Transactions sur ce compte.

**Soldes personnalisés** (`Store.settings.customBalances`, tableau de `{title, accounts}`) : jusqu'à 2 soldes additionnels, chacun calculé sur une sélection de comptes, affichés dans le bloc solde à côté du solde total. Titre vide = solde masqué. Inclus dans la sauvegarde `.json`/Gist, pas dans l'Excel (§7.2).

### 6.2 Transactions

**Bouton flottant unique** (bas d'écran, visible uniquement sur cet onglet, `#floatbar-wrap`) — 4 icônes :
- **Affichage** (`#btn-display`, icône `layers`) → `openDisplayMenu()`, sheet listant les 11 modes de regroupement (voir plus bas) ; sélection = fermeture immédiate + re-rendu. Pill cliquable (`#txn-view-indicator`, ex. "Jour ▾") qui rouvre la même sheet, à côté du bouton "Tout déplier/replier".
- **Loupe** (`#btn-search-toggle`) → bascule entre la rangée des 4 icônes (`#float-bar-actions`) et un champ de recherche plein cadre (`#float-bar-search`, bouton fermer `#btn-search-close`) dans le même espace. L'icône s'allume (`.active-filter`) si `Filters.search` est non vide, même champ replié.
- **Filtre** (`#btn-filter`) → sheet de filtres (§4), badge numérique = `Filters.activeCount()`.
- **Plus** (`#btn-add-txn`) → `openTxnNew()`, formulaire de saisie (voir plus bas).

**Regroupement** — 11 modes (`GROUPBY_LABEL`) : **Compte, Type de compte, Jour, Mois indiqué, Mois calculé, Groupe, Période, Catégorie, Personne, Groupe de personnes, Type**. "Mois indiqué" = mois de la date réelle ; "Mois calculé" = `AGG.getCalcMonthYM(t)` (§2.4).

**Groupe** et **Personne** affichent le libellé brut de la transaction (`dimDisplayValue`), pas le groupe de groupe/groupe de personnes résolu — cette résolution est portée par les dimensions **Période** et **Groupe de personnes**, séparées pour ne jamais mélanger niveau précis (groupe/personne) et niveau conteneur (période/groupe de personnes). De même, **Compte** regroupe par compte réel, **Type de compte** (§3) est un mode à part.

**Cartes de groupe — un seul rendu unifié pour tous les modes** (`computeTxnGroups`/`renderTxnGroupCards`) : chaque groupe (jour, mois, compte, groupe, catégorie ou personne) est une carte repliable `<details class="card txn-group-card">`, en-tête = libellé + nombre + total du groupe (revenus+dépenses, coloré), systématique dans tous les modes.
- Aucun groupe ouvert par défaut. L'état déplié/replié de chaque groupe se mémorise pendant la session (`APP.state.txnOpenGroups`, `Map` clé de groupe → booléen) et survit aux re-rendus déclenchés par "Afficher plus" ou le scroll infini — réinitialisé par `resetTxnPagination()` (changement de vue/filtre, retour sur l'onglet).
- **Bouton "Tout déplier"/"Tout replier"** (`#txn-expand-toggle`) : réglage global et persistant (`Store.settings.txnGroupsExpanded`, booléen, `false` par défaut). Sert de valeur par défaut pour tout groupe sans dépassement individuel (`openGroups.has(key) ? openGroups.get(key) : allExpanded`) ; cliquer efface tous les dépassements individuels en cours.
- **Chargement automatique au scroll** : un `IntersectionObserver` sur une sentinelle en bas de liste (`#txn-scroll-sentinel`, `setupTxnInfiniteScroll`) incrémente `APP.state.txnGroupsShown` (+20) à l'approche du viewport (`rootMargin:"600px"`). Le "Afficher plus (N restantes)" au sein d'un groupe déjà ouvert (au-delà de 10 transactions) reste un bouton manuel (`txnPerGroupShown`).

**Carte d'une transaction — 3 lignes, compactes** (`makeTxnItemEl`) :
- **Bandeau vertical à gauche** (`.txn-band`) : acronyme du compte (§3), coloré selon le type (vert=revenu, rouge=dépense, bleu=virement).
- **Ligne 1** (`.txn-title`, gras) : remarque, ou "Transfert X → Y", ou catégorie/"Revenu"/"Dépense" par défaut (`txnTitle`). Passe à la ligne si nécessaire (`white-space:normal;overflow-wrap:break-word`, pas de troncature) — une remarque longue s'affiche en entier, sur autant de lignes que nécessaire.
- **Ligne 2** (`.txn-sub`) : Catégorie · Personne (ou "compte → compte cible" pour un virement) — `txnLine2`.
- **Ligne 3** (`.txn-line`, la plus discrète) : Date courte · Groupe — `txnLine3`.
- **Montant** à droite, aligné en haut de la carte (pour rester cohérent quand la ligne 1 déborde sur plusieurs lignes).

Lignes 2 et 3 restent tronquées avec points de suspension — seule la ligne 1 (titre) peut s'étendre sur plusieurs lignes. Padding horizontal de la page réduit à 12px (au lieu des 20px standard de `.section` ailleurs), marge latérale de chaque transaction de 6px (au lieu de 8px). Le tap sur une carte ouvre la fiche transaction unifiée (ci-dessous).

**Fiche transaction (détail + édition unifiées)** — `openTxnDetail(id, navIds?)` et `openTxnNew()` convergent vers `renderTxnDetail(id, seed, navList)`. **Chaque champ se modifie au tap direct, sans mode "Modifier" séparé** : toucher la ligne Groupe ouvre le picker Groupe, toucher Montant ouvre un éditeur compact, etc.

- **Remarque en titre** : au-dessus du bloc "Détails", gras 19px, "(Sans titre)" en grisé si vide — tactile comme les autres champs (`openInlineEditor` en mode multiligne).
- **Bloc "Détails"** (une seule carte) : chaque champ y est une ligne tactile (`.txn-tap-row`, retour visuel `:active`, chevron `›`) — **Montant** (en premier, police 21px/900, coloré selon le signe), **Date**, **Type**, **Compte** (ou **Compte source**/**Compte cible** si virement, champs Catégorie/Personne masqués), **Catégorie**, **Personne**, **Groupe**.
- **Pickers réutilisés** pour Catégorie (`categoriePickerHierarchy()`, groupée par type Revenus/Dépenses, en-têtes non sélectionnables et sections dépliées par défaut), Personne (`AGG.personneHierarchy`, groupée par groupe de personnes), Groupe (`AGG.groupeHierarchyChrono`, tri chronologique, §2.1 bis), Compte/Compte cible (liste plate) ; Type via un picker à 3 options fixes avec `noEmpty:true`.
- **`openInlineEditor(opts)`** (réutilise `overlay-confirm`/`modal-confirm`, modale compacte) : édite Montant (`inputType:"number"`), Date (`inputType:"date"`) ou Remarque (`multiline:true`, `<textarea>`) via un simple champ + Annuler/Enregistrer.
- **Transaction existante — chaque champ s'enregistre immédiatement** (`commit(patch)`) : reconstruit la ligne complète (`DATA.rowToTxn`) et remplace l'entrée dans `Store.transactions`. L'id étant dérivé du contenu (`buildTxnId`), il change à chaque modification — suivi via `currentId` fermée sur la fiche ouverte (`APP.state.detailTxnId` mis à jour en parallèle). Sauvegarde asynchrone (`Store.saveTransactions`) + rafraîchissement de l'onglet courant après chaque champ.
- **Nouvelle transaction (y compris "Dupliquer")** : champs accumulés dans un brouillon local (`draft`, montant non défini au départ — "Toucher pour ajouter") puis validés via **Enregistrer** (date et montant > 0 requis, sinon toast ; anti-collision d'id par suffixe aléatoire si besoin).
- **Boutons du bas** : transaction existante → flèche précédente (`#td-nav-prev`, si contexte de navigation) + **Dupliquer** (`#td-duplicate`, icône seule) + **Supprimer** (`#td-delete`, icône seule) + flèche suivante (`#td-nav-next`) ; nouvelle transaction → **Annuler** + **Enregistrer**.
- **Dupliquer** : ouvre le flux "nouvelle transaction" (`renderTxnDetail(null, source)`, sans `navList`) pré-rempli avec tous les champs de la source, date d'origine incluse.

**Navigation précédent/suivant dans la fiche transaction** — disponible uniquement quand la fiche est ouverte depuis la liste principale : `makeTxnItemEl(t, navIds)` reçoit `navIds`, calculé une fois par rendu (`renderTransactions` → `flattenTxnGroupIds(arr)`) en aplatissant tous les groupes/éléments filtrés-et-triés dans leur ordre d'affichage, y compris groupes repliés et au-delà de la pagination — la séquence couvre l'intégralité du résultat filtré. Les fiches ouvertes depuis un autre onglet (Budget, Récurrent, Évolution, Santé) via `openTxnDetail(id)` sans second argument n'ont pas de navigation : les flèches sont absentes.
- **Instantané figé à l'ouverture** (`nav`, copie locale de `navList`, jamais recalculée depuis les filtres tant que la fiche reste affichée) : modifier un champ qui ferait sortir la transaction du filtre actif ne casse pas la navigation ni ne fait disparaître la transaction en cours d'édition. `nav[navIndex]` est mis à jour avec le nouvel id après chaque `commit()`. Cet instantané ne survit pas à la fermeture de la fiche : rouvrir une transaction repart d'un `navIds` recalculé sur le filtre à jour.
- **Trois déclencheurs équivalents** : boutons flèche (`#td-nav-prev`/`#td-nav-next`, grisés en début/fin de séquence), balayage tactile gauche/droite (`modal-detail.ontouchstart`/`.ontouchend`, seuil 60px avec déplacement horizontal ≥1,5× le vertical), et flèches clavier `ArrowLeft`/`ArrowRight` (écouteur global posé dans `buildShell`, actif seulement quand `APP.state.txnNavActive` est renseigné, qu'aucun champ n'a le focus, et qu'aucune autre fenêtre n'est ouverte par-dessus).
- **Nettoyage automatique** : `closeOverlay("overlay-detail")` réinitialise `txnNavActive` à `null` et efface les gestionnaires tactiles du modal — nécessaire car `modal-detail` est un conteneur **partagé** par de nombreuses autres fiches (réordonner les comptes, soldes personnalisés, écrans du Budget…).

**Confirmation avant suppression** (`APP.confirmAction(opts)`, modale générique réutilisable — `{title, text, confirmLabel, onConfirm}`, `text` supporte `\n`) : protège la suppression d'une transaction, d'un ajustement budgétaire (§6.3), et d'une règle groupe/personne obsolète (§2.4). `overlay-confirm` doit rester le **dernier élément du shell HTML** pour toujours s'afficher au-dessus.

### 6.3 Budget

**Deux modes** :
- **"Amort. + ajouts"** (`cfg.mode==="amorti"`) : amortissement + ajustements actifs.
- **"Ajouts"** (`cfg.mode==="brut_ajout"`) : ajustements actifs, pas de lissage des catégories.

Bouton séparé, à droite des deux boutons de mode (icône colonnes, `bg-detail-toggle-btn`) : bascule la **vue détail** (`cfg.detailView`, persisté). Activée, chaque montant (catégories, voyages, totaux de section, total de carte, "Total tous groupes") s'affiche en **3 colonnes** au lieu d'une valeur unique :
- Mode "Amort. + ajouts" : **Réel + ajout · Amorti · Total**.
- Mode "Ajouts" : **Brut · Ajout · Total**.

Chaque valeur des 3 colonnes est individuellement cliquable (classe `.bg-detail-cell`, wiring central dans `renderBudget`) et ouvre la feuille de détail correspondante, préfiltrée sur le type exact de la colonne (`filterKind`: `"reel"`, `"amorti"`, `"ajout"`, ou `"reel_ajout"` — filtre composé géré par `bgKindIncluded()`). Fonctionne aussi pour les voyages et pour "Total tous groupes". Voir §6.3.4 pour le calcul Réel/Amorti/Ajout côté salaire.

**Scope** : Perso (×1) ou Couple (×2, appliqué aux transactions réelles seulement — ajustements en montant brut non multiplié).

**Vue** : Total ou **Moyenne**. Divise les totaux affichés par le nombre de mois de la période affichée **compris dans la borne chronologique de l'amortissement** (`Store.settings.budget.startMonth` → `endMonth`, `amortBoundedMonths`) si une des deux bornes est renseignée, sinon par le nombre de mois de la période affichée complète — quel que soit le mode.

**Période affichée** (`periodPickerHtml`, `periodFrom`/`periodTo`) : chaque borne peut être laissée **"Non définie"** — bouton "✕" à côté d'un mois renseigné pour la vider, bouton texte "Non définie" pour lui redonner un sélecteur. Début non défini = tout l'historique disponible jusqu'à la fin ; fin non définie = tout ce qui suit le début, jusqu'au mois en cours ou à la dernière donnée connue si plus tardive. La borne manquante est résolue par `budgetDataMonthBounds()` (étendue mini/maxi parmi transactions, salaires saisis et ajustements, mois en cours inclus comme plancher haut) avant `monthsInRange`. Élargir la période affichée ne contourne aucune règle d'amortissement : la fenêtre d'amortissement, le bornage des groupes additionnels (`extraGroupEffectiveMonths`) et celui des salaires (§6.3.4) restent calculés à partir de `startMonth`/`endMonth`. Persisté dans `Store.settings.budget.periodFrom`/`periodTo` (peut valoir `null`).

L'onglet Budget affiche toujours les montants sans décimale, indépendamment du réglage global (§5).

**Quatre boutons toujours visibles, sur une seule ligne**, ouvrent chacun un pop-up dédié :
- **Amortissement** → mois de début, période de lissage (calculée automatiquement), catégories amorties (pilules One UI), case **"Ne pas distinguer les voyages"** (§6.3.2 bis) ; sous-libellé "Inactif (mode Ajouts)" quand `cfg.mode !== "amorti"`.
- **Ajustements** → liste des ajustements de la période ; sous-libellé "N sur la période" ou "Aucun" (formulaire détaillé en §6.3.6).
- **Comparaison** → réglages de comparaison ; sous-libellé la période comparée, "Budget prédéfini", ou "Désactivée".
- **Comptes exclus** → pilules One UI cliquables, une par compte ; sous-libellé "N exclu(s)" ou "Aucun".

#### 6.3.1 En-tête de colonnes figé ("Total tous groupes")
Quand une comparaison est active OU que la vue détail est activée, un en-tête de colonnes (`compareHeaderRowHtml()` ou `detailHeaderRowHtml(cfg.mode)`) est affiché au-dessus de la carte "Total tous groupes", dans sa propre petite carte flottante, en `position:sticky;top:0;z-index:10`, directement enfant de la section qui contient aussi toutes les cartes de groupes. Sa boîte englobante pour le calcul du sticky est donc toute la section, pas seulement la petite carte de total — il reste figé en haut de l'écran pendant tout le défilement de la liste des cartes de groupes.

#### 6.3.2 En-tête de carte de groupe
Nom du groupe + chevron cliquable → sélecteur pour changer le groupe occupant cette carte. Icône avion ✈️ (§6.3.2 bis) pour regrouper les voyages de cette carte dans ses catégories. Pour un groupe additionnel uniquement, bouton **Retirer** (✕) supplémentaire. Les groupes additionnels suivent toujours les mêmes mois que le principal (`extraGroupEffectiveMonths` = intersection période affichée / borne d'amortissement).

**Carte du groupe principal et d'un groupe additionnel** : structure identique — Total + listes **Revenus**, **Dépenses** et **Voyages**, chaque ligne cliquable → détail des transactions (ou 3 colonnes cliquables si vue détail active). Sections Revenus/Dépenses/Voyages repliables (`<details>`) mais partagent un seul état global (`allSectionsOpen`).

#### 6.3.2 bis Regrouper les voyages dans les catégories ("Ne pas distinguer les voyages")
Par défaut, un groupe de type Voyage (`getGroupKind`, §2.1) est affiché à part : ses transactions alimentent le bloc "Voyages" de sa carte (une ligne par occurrence de voyage), jamais les listes Revenus/Dépenses par catégorie. Deux réglages, cumulables (`voyagesMergedForGroup(groupName)` = l'un OU l'autre), désactivent cette distinction et fondent ses coûts/revenus dans les catégories du groupe, comme une transaction "vie quotidienne" ordinaire :
- **Réglage global** (`Store.settings.budget.mergeVoyages`, case dans la fiche Amortissement) : s'applique à tous les groupes de la carte.
- **Réglage par groupe** (`Store.settings.budget.mergeVoyagesGroups`, tableau de noms) : icône avion dans l'en-tête de chaque carte (§6.3.2), colorée quand active pour cette carte précise. Indépendant d'une carte à l'autre.

**Le regroupement ne change jamais le mois de comptage d'une transaction, ni son éligibilité à l'amortissement — seule la ligne sous laquelle elle apparaît change.** La résolution du mois d'une transaction de groupe Voyage (`voyageTxnMonths()`, fonction centralisée partagée par `computeGroupBudget`, `openCategoryTxnSheet` et `collectGroupTypeItemsDetailed`, pour que total et détail cliquable restent toujours cohérents) garde la même priorité, fusionné ou non : "Mois voyage" forcé (§2.4) d'abord, sinon date corrigée (transform, §1.3), sinon date réelle. Une transaction de groupe Voyage n'est jamais amortie, fusionnée ou non — le regroupement change uniquement la case de destination, jamais le calcul du montant ni son mois.

#### 6.3.3 Détail d'une catégorie (tap sur une ligne "Catégories")
- Cas spécial **catégorie "Salaire"** (nom insensible casse/accents) : tap ouvre l'éditeur de saisie manuelle (§6.3.4) au lieu du détail habituel.
- Badges **Réel / Amorti / Ajout** en haut (totaux) — cliquables, servent de filtre sur la liste en dessous.
- Montant affiché par ligne = montant utilisé par le budget, cohérent avec les totaux de la carte.
- Chaque transaction "Réelle" a un bouton **"Amortir"** : ouvre le formulaire d'ajustement pré-rempli pour convertir cette transaction précise en dépense amortie.
- Montants "lissés sur l'année" (yearsplit) affichés au prorata de la période ; libellé court "XX/12 mois".

#### 6.3.4 Éditeur "Salaire" — décomposition Brut/Ajout
Remplace le détail habituel pour la catégorie dont le nom normalisé est "salaire". Une ligne par mois de la période, 3 colonnes : **Moi** (saisie), **Ma femme** (saisie), **Total** (calculé). Stocké dans `Store.settings.budget.salaryEntries[mois] = {moi, femme}`.

**Logique Brut/Ajout** (`salaryMonthBreakdown`) :
- Le salaire de **"Ma femme"** est toujours entièrement "Brut" — sa valeur déclarée n'est jamais comparée à des transactions.
- **"Moi"** : le montant des transactions réelles de la catégorie ce mois-là est toujours le "Brut". Si une saisie manuelle est renseignée, elle représente le **salaire déclaré** ; l'écart entre valeur déclarée et montant des transactions devient l'**"Ajout"** (positif ou négatif). Sans saisie manuelle, "Moi" = Brut seul (Ajout = 0).
- **Total** : Couple = (Brut+Ajout de Moi) + Ma femme (saisi ou 0) ; Perso = Moi uniquement.
- **Jamais compté hors de la borne chronologique de l'amortissement** (`amortBoundedMonths`, mêmes bornes `startMonth`/`endMonth` que les catégories amorties et les groupes additionnels) : un mois de la période affichée hors de cette fenêtre ne contribue ni au Brut ni à l'Ajout du salaire.

Cette décomposition alimente directement `o.realIncome`/`o.adjIncome` de la catégorie Salaire dans `computeGroupBudget`, compatible avec la vue détail (§6.3) : la ligne Salaire affiche Réel+Ajout/Amorti/Total (Amorti toujours à 0, le salaire n'étant jamais amorti).

#### 6.3.5 Comparaison
`Store.settings.budget.compare = {enabled, mode:"period"|"predefined", periodFrom, periodTo, predefined:{}}`. Une fois activée, l'affichage bascule en tableau à 3 colonnes (Résultat · Comparé · Variation). Les totaux sont toujours la somme des lignes affichées en dessous. Voyages jamais détaillés voyage par voyage en vue comparaison — une seule ligne "Total voyages", toujours affichée même à 0 €/0 €.
- **Mode "Autre période"** : le scénario comparé applique exactement les mêmes réglages que le principal, seule la période change.
- **Mode "Budget prédéfini"** : montants mensuels saisis à la main, un par catégorie de chaque groupe et un par groupe pour l'ensemble de ses voyages.

**Catégories alimentées seulement côté comparé** (`compareOnlyCategories()`) : une catégorie sans transaction dans le scénario principal sur la période affichée, mais alimentée côté comparé, n'est jamais masquée — elle s'affiche avec un montant principal à 0 €. Sans cela, sa valeur comparée serait invisible et absente des totaux Revenus/Dépenses/Total. Type (Revenus/Dépenses) déterminé par le classement global de la catégorie (`classifyCategoryTypes`), jamais par le signe de la valeur comparée.

**Budget prédéfini — toujours un budget réel** : chaque montant prédéfini saisi (`Store.settings.budget.compare.predefined`, objet clé→montant, `predefinedCatKey`/`predefinedVoyageKey`) est considéré comme un montant réel, quelle que soit la vue affichée — plus de choix par ligne (l'ancien sélecteur Amorti/Réel/Les deux et le champ `compare.predefinedVue` ont été supprimés). En vue "Ajouts", le montant est utilisé tel quel. En vue "Amort. + ajouts", il est complété par la part d'amortissement correspondante (`predefinedRealValue`), pour rester cohérent avec le reste du budget.

La "part d'amortissement" d'une ligne n'est pas un chiffre saisi à part : c'est la contribution `amortIncome`/`amortExpense` calculée pour cette même catégorie (ou pour la ligne "Ajustements" côté voyages, seule à pouvoir en porter une) par `computeGroupBudget`, recalculée en forçant le mode "amorti" même quand la vue affichée est "Ajouts" (`opts.mode`, résultat mis en cache par groupe dans `computeCompareContext.getAmortResult`, jamais appliqué à un groupe additionnel qui reste toujours en brut). Mise à l'échelle Total/Moyenne identique au montant prédéfini lui-même.

*Exemple* : un montant prédéfini de 1000 € de dépense, avec 700 € d'amortissement calculés sur cette même ligne, affiche **1000 €** en vue "Ajouts" et **1700 €** en vue "Amort. + ajouts".

**Clic sur la cellule "Comparé"** : chaque valeur de la colonne "Comparé" (catégorie, Total voyages, Total de carte, Total Revenus/Dépenses, grand total tous groupes) est individuellement cliquable (classe `.bg-cmp-cell`, wiring central dans `renderBudget`) et ouvre exactement la même feuille de détail que le scénario principal, mais avec les montants de la **période comparée**, via `opts.periodOverride` transmis à `openCategoryTxnSheet`/`openVoyageTxnSheet`/`openGroupTypeTxnSheet`. En mode "Budget prédéfini", le clic ouvre l'éditeur du budget prédéfini. Les cellules "Résultat" et "Variation" ne sont jamais individuellement cliquables : le clic remonte au gestionnaire de la ligne entière.

#### 6.3.6 Ajustements manuels — formulaire (`openAdjustmentForm`)
Champs : mois, **groupe** (voir ci-dessous), catégorie (vide = ajustement global au groupe, imputé en `Ajustements` parmi les voyages, §6.3.2), montant (négatif = dépense), remarque, scope de saisie (Perso ×1 / Couple ×2 / Les deux, §6.3), case "n'apparaît que dans cette vue" (`montantViewOnly`), case "Amortir ce montant sur toute la période" (`amorti`).

**Champ Groupe** : bouton (`PICKER.fieldPickerBtnHtml`) ouvrant le même picker hiérarchique que la fiche transaction (§6.2), via `adjustmentGroupHierarchy()` — `AGG.groupeHierarchyChrono` (§2.1 bis) : groupes de groupe sélectionnables en bloc (en-tête de section), triés du plus récent au plus ancien, chaque groupe précis sélectionnable individuellement en dessous (trié du plus ancien au plus récent), groupes non classés à part (triés alpha). Un ajustement peut cibler **n'importe quel groupe**, précis ou de groupe, exactement comme un groupe principal/additionnel du Budget (§6.3.2).

Chaque groupe précis (jamais les en-têtes) est précédé d'une icône Voyage ✈️/Vie quotidienne 🏠 (`getGroupKind`, §2.1) via `h.kindFn` de `renderPickerList` (§12) — pour voir dans quelle case budgétaire (catégories "vie quotidienne" vs bloc "Voyages") l'ajustement ira une fois enregistré.

**Correspondance avec une carte du Budget, avec repli hiérarchique** (`adjustmentGroupMatches(adj.groupe, superGroupeName)`, module Budget) : un ajustement compte pour une carte (`cfg.mainGroup`/un élément de `cfg.extraGroups`) si son champ Groupe correspond exactement, ou correspond au nom de base, **ou si son groupe précis résout vers le groupe de groupe de la carte** (`AGG.labelResolvesToSuperGroupe`, via `groupAliasRules`) — même principe que `groupSelectionMatches` pour les transactions réelles. Exemple : une carte configurée sur le groupe de groupe "Paris" reçoit automatiquement un ajustement posé sur son groupe enfant précis "Paris 2025". **Pas de repli dans l'autre sens** : un ajustement posé directement sur un groupe de groupe ne redescend jamais sur une carte configurée sur l'un de ses groupes enfants précis (ambigu s'il existe plusieurs cartes enfants) — le champ Groupe doit alors correspondre exactement au réglage de cette carte précise.

### 6.4 Évolution
Vue tabulaire de l'évolution d'une dimension dans le temps, avec drill-down année → mois, affichage cumulé ou par période (§6.4.2 bis) et ouverture directe de la liste de transactions depuis n'importe quel montant.

#### 6.4.1 Dimensions et sections
Sélecteur de vue à 4 boutons (`Store.settings.evolution.view`, persisté) : **Groupe**, **Catégorie**, **Personne**, **Compte**.
- **Groupe** : sections = groupe de groupe résolu (§2.1), items = libellé complet du groupe.
- **Catégorie** : sections = Revenus/Dépenses (`classifyCategoryTypes`, avec override), items = catégorie.
- **Personne** : sections = groupe de personnes résolu, items = personne.
- **Compte** : sections = type de compte (§3), items = compte. Seule dimension où les **virements sont inclus** (les 3 autres les excluent) : un virement contribue `-montant` au compte source et `+montant` au compte cible. Le total général ("Total" en haut, §6.4.4) neutralise toujours les virements internes.

Moteur d'agrégation : `AGG.buildEvolutionIndex(dimension, monthBasis)`, retourne `{sections:[{key,label,items:[{key,label,values:Map("YYYY-MM"→montant)}]}]}`. Sections/items triés par montant total décroissant.

#### 6.4.2 Base mensuelle : réel vs transformé
Sélecteur à 2 boutons (`Store.settings.evolution.monthBasis`, persisté) :
- **Mois réel** : chaque transaction comptée sur sa date réelle (`t.date`), sans transformation.
- **Mois transformé** : applique `t.transform` (§1.3) — une transaction "yearsplit" répartie également sur les 12 mois de son année cible, comme dans le moteur Budget.

#### 6.4.2 bis Mode cumulé
Troisième sélecteur à 2 boutons ("Normal"/"Cumulé", `Store.settings.evolution.cumulative`, persisté **localement** — absent de `buildParamsPayload`, §7.2, comme `theme`).

En mode Cumulé, chaque colonne (année ou mois) affiche la somme de toutes les valeurs depuis le début des données jusqu'à cette période incluse — comparaison lexicographique sur les clés `"YYYY-MM"` : `AGG.itemCumulativeYearValue`/`sectionCumulativeYearValue` (borne = décembre de l'année), `AGG.itemCumulativeMonthValue`/`sectionCumulativeMonthValue` (borne = le mois exact). La cumulation reste continue au travers d'un drill-down.

**La colonne "Total" n'est jamais recalculée par ce mode** — toujours la somme brute (`itemTotalValue`/`sectionTotalValue`), identique dans les deux modes.

Le clic sur un montant en mode Cumulé ouvre la même feuille de détail qu'en mode Normal (§6.4.3), avec les transactions cumulées correspondantes (toute transaction dont la période est ≤ la colonne cliquée) et un intitulé adapté ("Jusqu'à *période* (cumulé)"). Un clic sur "Total" reste "Toute la période" dans les deux modes.

#### 6.4.3 Colonnes et drill-down
Tableau HTML (`<table class="evo-table">`, pas de grille CSS) pour un comportement de colonnes figées fiable : les 2 premières colonnes (**Libellé**, **Total**) sont en `position:sticky;left:...`, largeurs fixes en pixels, `box-sizing:border-box` sur toutes les cellules (indispensable, sinon le padding s'ajoute aux `min/max-width` et désaligne l'offset `left` de la 2ᵉ colonne figée). Les colonnes suivantes défilent horizontalement : une par **année** par défaut, ou une par **mois** (Jan.→Déc.) après clic sur une année.

**Largeur et défilement** : padding latéral réduit à 8px, colonnes Total/valeur en largeur compacte uniforme de 74px. Défilement horizontal à la main au-delà : `U.enableDragScroll(el)` (§12) attache un glisser-déposer souris en plus du défilement tactile natif (`overflow-x:auto`) — un seuil de déplacement distingue un clic (ouvre le détail) d'un glissé (fait défiler). Repris par l'onglet Récurrent (§6.5) et la vue mensuelle du Budget.

- **Clic sur une année** (`.evo-year-header`) : `drillYear` passe à cette année, le tableau réaffiche 12 colonnes mensuelles, avec un bouton "Retour aux années". La colonne "Total" affiche alors le total de l'année sélectionnée.
- **Clic sur un montant** (`.evo-val-cell`) : ouvre `openEvolutionTxnSheet(dimension, opts)`, filtre `Store.transactions` par dimension (`txnMatchesDimTarget`) et période (`txnMatchesPeriod`) puis affiche la liste triée par date, avec un badge Total.
- **Clic sur le libellé d'une section** : replie/déplie ses items.

Vue **Groupe** uniquement : sections et items triés chronologiquement (§2.1 bis), pas par montant.

#### 6.4.4 Ligne "Total" générale
Première ligne du tableau (toujours visible, jamais repliable), somme de toutes les sections pour chaque colonne. Cliquable (`data-global="1"`) — ouvre la liste de toutes les transactions de la dimension pour la période cliquée.

#### 6.4.5 Fusion de sections
Une fusion (`{name, members:[clés de section]}`, stockée dans `Store.settings.evolution.merges[dimension]`) se configure uniquement via la sauvegarde `.json` ou la synchronisation Gist (§7.2) — pas de création depuis l'interface, pas de feuille Excel dédiée.

L'onglet Évolution se limite à :
- **Afficher** les fusions actives sous forme de chips sous les sélecteurs de vue.
- **Supprimer** une fusion individuellement (bouton × sur sa chip, défusion immédiate et persistée localement).

À l'affichage, `buildEvolutionIndex` fusionne les sections membres en une seule section synthétique (`__merged:true`, items préfixés en interne pour éviter les collisions de clé, affichés sous leur libellé d'origine).

Inclus dans la sauvegarde `.json`/Gist, fusionné par dédoublonnage à l'import (§7.2) — contrairement à la vue actuellement affichée (`evolution.view`/`monthBasis`/`cumulative`), jamais modifiée par un import "Ajouter".

---

### 6.5 Récurrent
Isole les transactions dont la Remarque porte un suffixe de date transformée (`AA-MM` ou `AAAA`, §1.3, `t.transform.mode !== "same"`) — celles dont le **mois calculé** diffère du mois réel — regroupées par catégorie, pour suivre des dépenses/revenus récurrents indépendamment de leur date réelle d'enregistrement. Même format visuel que l'onglet Évolution (tableau à colonne Libellé figée, sections repliables, montants cliquables, défilement compact §6.4.3).

**Moteur** : `AGG.buildRecurringIndex(periodFrom, periodTo)` — sections = catégorie (`t.categorie || "(Non renseigné)"`), items = **nom de la transaction sans le suffixe de date** (`U.baseRemarque(remarque)`, symétrique de `computeTransform`). Items triés alphabétiquement au sein de chaque section ; sections triées par montant absolu décroissant.

**Colonnes** : Total de la période, puis une colonne par mois. Une entrée `AA-MM` alimente son mois précis (`item.values`) ; une entrée `AAAA` (répartition sur l'année) n'est **pas** éclatée sur 12 mois ici — elle reste groupée par année entière (`item.yearValues`) et s'affiche dans une **colonne spéciale juste après décembre** de cette année (`AGG.recurringColumns`). La colonne Total **n'est pas figée** (contrairement à Évolution) : simplement la première d'une séquence qui défile — seule la colonne Libellé reste sticky.

**Filtre de période** (`Store.settings.recurrent.periodFrom/periodTo`, deux `<input type="month">`) : porte uniquement sur le **mois calculé** des entrées `AA-MM`, ou sur décembre de l'année pour une entrée `AAAA` — jamais sur le mois réel. Par défaut, la période couvre tout l'historique disponible (`ensureRecurrentPeriod`, calculée une fois puis persistée).

**Tout déplier/replier** (`#rec-expand-toggle`) : même mécanisme que Transactions (§6.2) — réglage global persistant `Store.settings.recurrent.groupsExpanded`.

**Clic sur un montant** : ouvre `openRecurringTxnSheet(opts)` (même esprit que `openEvolutionTxnSheet`) ; le filtrage (`txnMatchesRecurringTarget`) distingue trois cas — colonne d'un mois précis, colonne année spéciale, ou colonne Total. Chaque ligne ouvre à son tour la fiche transaction standard (`openTxnDetail`, §6.2).

---

### 6.6 Santé
Onglet secondaire, indépendant des Transactions (données, filtres et listes propres) mais reprenant la même présentation — bouton flottant, style de carte, fiche détail par tap direct. Accessible via le bouton "⋯" (`#nav-more`, dernière position de la barre de navigation) qui ouvre une sheet (`overlay-more`/`sheet-more`, `openMoreMenu()`) listant les onglets secondaires — `MORE_TABS = [{key:"revenus",...}, {key:"sante",...}]`. `showFloat` (§6.2) inclut `"sante"` en plus de `"transactions"` (pas `"revenus"`, §6.7) ; les 4 boutons du bouton flottant se dispatchent selon `currentTab`.

**Modèle** (`Store.santeItems`, un item = un soin) :
```
{ id, dateSoin, beneficiaire, devise, montantEur, montantDevise,
  medecin, typePrestation, statut, statutDate,
  rembTiersMontant, rembTiersDate, rembSecuMontant, rembSecuDate, rembMutuelleMontant, rembMutuelleDate,
  groupe, mutuelle, partMoi, termine }
```
`montantDevise` n'existe que si `devise !== "€"` (sinon `null`) — purement informatif à côté de `montantEur`, aucune conversion/taux de change automatique. `id` = hash de contenu (`buildSanteId`, même principe que `buildTxnId`, §1.2) sur `dateSoin/beneficiaire/typePrestation/medecin/montantEur/groupe/statut` — **`termine` en est exclu** : basculer ce statut ne change ni l'id ni la déduplication à l'import.

**"Ma part"** (`computeMaPart(it)`) = `montantEur × partMoi ÷ 100`, arrondi à la **2ᵉ décimale inférieure** (`Math.floor`, jamais un arrondi standard) — affichée à côté du pourcentage dans la fiche (ligne "Part moi") et en totaux filtrés dans l'en-tête de liste.

**Statut "Terminé"** (`termine`, booléen) : bouton pilule en haut de la fiche détail (`#sd-termine`), bascule immédiate avec sauvegarde. Un soin terminé s'affiche grisé (`opacity:0.6`) avec "✓ Terminé" en ligne 3 de sa carte. Filtrable (2 puces fixes) et regroupable (mode d'affichage dédié, non-terminés triés avant terminés).

**Listes personnalisables** (`Store.settings.sante.lists`, 6 clés : `beneficiaires, devises, medecins, prestations, statuts, mutuelles` — `devises` pré-rempli `["€","$","£","CHF"]` à la première utilisation si vide) : chaque champ utilise `PICKER.openPicker({..., allowNew:true})`, toute valeur saisie à la volée est ajoutée automatiquement (`addToList`) — gérables aussi depuis le bouton réglages. **Le champ Groupe fait exception** : il pioche directement dans la hiérarchie des groupes des Transactions (`AGG.groupeHierarchyChrono`), présentée comme le filtre Groupe habituel (§2.1 bis, §4) — texte libre accepté aussi.

Sélectionner un **Statut** horodate automatiquement `statutDate` à aujourd'hui ; modifiable ensuite via la ligne "Mise à jour le".

**Liste et filtres** (`window.__SF_SANTE_FILTERS`, dimensions `beneficiaires/prestations/statuts/devises/groupes` en `Set` + `termine` + `search`, persistés dans `Store.settings.santeFilters`) : sheet de filtre à **page unique** (pas le menu en deux étapes des Transactions, §4), réutilisant `overlay-filter`/`sheet-filter`. Le filtre Groupe affiche les mêmes comptages que le filtre Groupe des Transactions (sur `Store.transactions`, pas sur les soins) — un compte affiché peut donc sembler élevé par rapport au nombre de soins réellement rattachés.

**Affichage** (`openSanteDisplayMenu`, réutilise `overlay-dispmenu`/`sheet-dispmenu`) : regroupement par Mois (défaut), Bénéficiaire, Statut, Groupe, Médecin/hôpital ou Terminé (`SANTE_GROUPBY_LABEL`) — cartes de groupe toujours dépliées, pas de pagination (volume de soins nettement plus faible que les transactions). En-tête de liste : nombre de soins + montant total filtré, puis Remboursé/Reste à charge/Ma part (`computeSanteTotals`), puis le pill "Affichage".

**Fiche soin** (`renderSanteDetail(id, seed)`, réutilise `overlay-detail`/`modal-detail` comme la fiche transaction, §6.2) : bouton Terminé + fermer en haut, titre "Type de prestation" tactile, puis cartes **Soin** (Montant €, Devise, Montant devise si ≠ €, Date du soin, Bénéficiaire, Médecin/hôpital, Groupe), **Statut** (Statut + date de mise à jour) et **Remboursements** (Tiers payant/Sécu/Mutuelle, chacun ouvre `openReimbEditor`) et **Mutuelle & part** (Mutuelle, Part moi — `openPartMoiEditor`, pilules rapides 0/50/100 % + saisie manuelle). Nouveau soin : brouillon local, "Enregistrer" exige date + montant > 0 ; soin existant : **Dupliquer**/**Supprimer** (confirmation via `APP.confirmAction`).

**Réglages Santé** (bouton dédié `#sante-settings-btn`, propre à Santé) : `window.__SF_SANTE_SETTINGS`, sheet propre (`overlay-santesettings`) avec sous-écrans — menu listant les 6 listes personnalisables (ajout/suppression simple, `renderSanteListEditor`) et deux actions Excel dédiées à l'onglet :
- **Export** (`exportSanteExcel`) : une feuille "Santé" dédiée, 19 colonnes (`SANTE_XLSX_HEADER`, dont "Terminé" en `Oui`/`Non`), dates au format Excel natif `dd/mm/yyyy`. Construction des lignes factorisée dans `buildSanteExportRows()`, réutilisée par l'export Excel principal (§8) quand la case "Santé" y est cochée.
- **Import** (`rowToSanteItem`) : `.xlsx`/`.xls`/`.csv`, en-têtes reconnus par nom normalisé, montants tolérants à la virgule décimale, "Terminé" reconnu depuis `Oui/Yes/True/1/x`. **Réconciliation casse/accents silencieuse** (même principe que §7.1 bis) : Groupe est aligné sur un groupe déjà utilisé par une transaction existante, et Bénéficiaire/Devise/Médecin/Type de prestation/Statut/Mutuelle sont alignés sur une valeur déjà présente dans la liste personnalisable correspondante (`U.alignToCanon`/`U.transactionFieldCanon`/`U.valueListCanon`) — appliquée avant le calcul de l'id (`buildSanteId`), pour qu'une simple faute de casse ne crée ni entrée fantôme dans une liste déroulante, ni doublon du soin à un réimport ultérieur. Fusion (dédup par id, comme Transactions) ou remplacement complet ; les valeurs réellement nouvelles sont ajoutées automatiquement aux 6 listes personnalisables (`registerNewListValuesFromItems`). `rowToSanteItem` est aussi utilisée par le flux d'import principal (§7) pour détecter et lire une feuille "Santé" au sein d'un classeur Transactions/Paramètres.

**Stockage et sauvegarde** : `Store.santeItems` en IndexedDB (clé `sante_items`, repli `localStorage` `sf_sante_items_v1`, §9) ; `Store.settings.sante`/`Store.settings.santeFilters` dans le même `localStorage` que le reste des réglages. Inclus dans la sauvegarde `.json` complète (§8) et dans le payload de synchronisation Gist (§10, remplacement complet, jamais de fusion) et, si la case dédiée est cochée, dans l'export Excel multi-feuilles principal. "Réinitialiser toutes les données" (Menu ☰) efface aussi les soins.

### 6.7 Revenus

Onglet secondaire (icône `pie`), accessible comme Santé via le bouton "⋯" (§6.6). Répartit les revenus entre **Moi** et **Ma femme** (clés internes `moi`/`conjoint`, libellés `REV_PERSON_LABEL`) avec des corrections de transfert entre les deux — pas de bouton flottant.

#### 6.7.1 Modèle
```
Store.settings.revenus = {
  periodFrom, periodTo,     // "AAAA-MM", filtre de période (mois calculé, cf. 6.7.5)
  groupFilter: [],          // noms de groupe/groupe de groupe (cf. 6.7.5)
  groupsExpanded,           // "Tout déplier" persistant, cf. 6.7.4
  moi:      { lines: [...lignes], transferOut: [...lignes], transferIn: [...lignes] },
  conjoint: { même forme }
}
```
Pas de regroupement par catégorie : `lines` est une liste **à plat**, directement sous "Avant transfert" (§6.7.2). Une **ligne** (`{id, type, label, sign, ...}`) est de l'un de ces trois types :
- **`manual`** : `values = {"AAAA-MM": montant}`, saisie libre.
- **`recurring`** : lit `AGG.buildRecurringIndex` (moteur de l'onglet Récurrent, §6.5) via `line.recurCategorie` + `line.recurLabel`, portés par la ligne elle-même — permet à une ligne récurrente de vivre aussi bien dans `lines` que dans un bucket de transfert (§6.7.3).
- **`budgetSalary`** : ne stocke aucun montant — lit/écrit **en direct** `Store.settings.budget.salaryEntries[AAAA-MM].femme` (le champ "Ma femme" du salaire, §6.3) : synchronisation bidirectionnelle automatique avec l'onglet Budget. Proposée uniquement sous **Ma femme** (`hasBudgetSalaryLine` limite à une seule instance), jamais sous Moi.

`sign` (`+1`/`-1`, défaut `+1`) : multiplicateur appliqué aux deux points d'entrée de lecture (`revLineColValue`/`revLineTotalValue`) — sert à une ligne récurrente dont le montant, tel qu'enregistré côté transactions, doit compter en sens inverse ici, sans toucher aux transactions sources. Réglable pour les trois types de ligne (bouton `±`, §6.7.6). Aucun signe structurel n'est imposé par le bucket (§6.7.3) : une ligne de `transferOut` prend `-1` par défaut à sa création, une ligne de `transferIn` prend `+1` — valeur de départ modifiable comme n'importe quelle autre ligne.

#### 6.7.2 Trois blocs et hiérarchie visuelle
Par personne, trois totaux dans cet ordre : **Avant transfert** (somme de `lines`, à plat) → **Transfert** (addition pure **Vers + Depuis**, jamais une soustraction — §6.7.3, coloré selon le signe) → total final de la personne (porté par la **ligne Personne elle-même**).

Format compact façon Récurrent (§6.5), sans cascade d'indentation : la ligne Personne (0 indentation, `.rev-person-row`) est le seul niveau distinct — "Avant transfert", "Transfert" et ses deux sous-blocs "Vers X"/"Depuis X" partagent tous la même indentation (22px, `.rev-cat-row`) ; les lignes de `lines` apparaissent directement sous "Avant transfert", à cette même indentation. Les libellés longs passent à la ligne (`white-space:normal` sur `.evo-col-label`, §12) — partagé par Évolution/Récurrent/Budget.

Repliage à deux `Set` module-locaux non persistés (même convention qu'Évolution/Récurrent) : `revCollapsedPersons` (clé = `moi`/`conjoint`) et `revCollapsedSections` (clés `"perso::avant"`, `"perso::transfert"`, `"perso::transfert::transferOut|transferIn"`). Toute nouvelle sorte de ligne repliable doit porter la classe `.hidden` sur son `<tr>` (règle CSS générique `.evo-table tr.hidden{display:none;}`, §12).

#### 6.7.3 Transferts — addition signée, miroir automatique
`transferOut`/`transferIn` sont deux buckets **fixes** par personne (ni renommables ni supprimables) : `moi.transferOut` (argent envoyé) est le miroir exact de `conjoint.transferIn` (argent reçu), et vice-versa (`revMirrorOf`). Les deux lignes miroir partagent le **même `id`** — pas de table de correspondance séparée.

**"Transfert" = `transferOut` + `transferIn`, addition simple, jamais une soustraction dans le code** (`revPersonTransferColValue`/`TotalValue`) : aucun signe structurel n'est imposé par le bucket lui-même — chaque ligne compte pour sa valeur affichée telle quelle (montant × son propre `sign`, §6.7.1). C'est le signe de la ligne, pas une convention cachée de calcul, qui détermine si elle réduit ou augmente le total.

**Le miroir recopie le signe INVERSÉ, pas à l'identique** (`revSyncTransferLineToMirror` : `mLine.sign = -revLineSign(line)`) — libellé, montants (ou lien récurrent) et type sont recopiés tels quels, seul le signe s'inverse d'un côté à l'autre. C'est ce qui fait qu'un même transfert réduit toujours le total de qui l'envoie et augmente celui de qui le reçoit, quel que soit le signe choisi sur la ligne d'origine. Toute création/renommage/changement de signe/édition de montant sur une ligne est répercutée sur sa miroir ; la suppression (`revRemoveTransferLineAndMirror`) retire les deux à la fois, avec confirmation nommant l'autre personne.

Une ligne de transfert peut être manuelle **ou** récurrente (jamais salaire synchronisé) — le choix récurrent passe par `pickRecurringTransaction(onPick)`, partagé avec l'ajout d'une ligne "Avant transfert" (§6.7.6) : la fonction demande toujours d'abord la catégorie récurrente source, puis l'élément.

#### 6.7.4 Colonnes, drill-down et édition
Même principe qu'Évolution (§6.4.3) : Total + une colonne par **année** de la période configurée (`revYearsList` — toutes les années de `[periodFrom, periodTo]`, y compris sans données, contrairement à Évolution), clic sur une année pour zoomer sur ses 12 mois, plus une colonne "lump" spéciale après décembre pour une année comportant une ligne récurrente à suffixe `AAAA` (même convention que Récurrent, §6.5). "Tout déplier/replier" (`#rev-expand-toggle`) : réglage global persistant (`groupsExpanded`).

Cellules éditables (`rev-editable`, soulignement en pointillés) : uniquement lignes manuelles/salaire synchronisé, jamais récurrentes. Clic sur un **mois** (vue zoomée) → `openRevValueEditor`, éditeur compact à un champ (`PICKER.openInlineEditor`). Clic sur une **année** (vue non zoomée) → `openRevYearEditor`, modale listant les 12 mois avec un champ chacun (`overlay-confirm`/`modal-confirm`) — remplit plusieurs mois d'un coup. Les deux écritures répercutent le miroir si la ligne éditée vit dans `transferOut`/`transferIn`.

#### 6.7.5 Filtres
**Période** (`periodFrom`/`periodTo`, deux `<input type="month">`) : même mécanique que Récurrent (`ensureRevenusPeriod`, période la plus large couvrant transactions récurrentes + montants manuels déjà saisis + salaire Ma femme déjà renseigné côté Budget).

**Groupe** (`openRevFilterSheet`) : groupé par groupe de groupe via `AGG.groupeHierarchyChrono`, ordre chronologique (§2.1 bis), sélectionnable au niveau du groupe de groupe (en-tête) ou d'un groupe précis, avec recherche. Implémentation propre à l'onglet, mêmes classes CSS (`.groupe-section`/`.groupe-section-head`/`.groupe-child-item`). Stocké dans `groupFilter` (tableau de noms, mélange groupe de base et groupe de groupe). Ne s'applique qu'aux lignes récurrentes : `AGG.buildRecurringIndex(periodFrom, periodTo, groupFilter)` teste `t.groupeBase` **et** `getGroupeDisplay(t)` — même sémantique à deux niveaux que le filtre Groupe de Transactions (§2.3).

#### 6.7.6 Gestion (sheet "Gérer mes revenus")
`overlay-revmgr`/`sheet-revmgr` (`openRevManager`) : bascule Moi/Ma femme, puis une seule carte "Avant transfert" (`avantLinesCardHtml`) listant toutes les lignes de la personne à plat (par ligne : signe `±` / réordonner ↕ / renommer ✎ [manuel uniquement] / supprimer) avec "+ Ajouter une ligne" (choix Manuel/Récurrent/**Salaire synchronisé (Budget)** si sous Ma femme et pas déjà présent). Puis deux cartes fixes "Transferts" (Vers X / Depuis X, jamais réordonnables entre elles) avec leurs propres lignes (mêmes actions, mirées, sans option Salaire synchronisé) et leur propre "+ Ajouter une ligne" (`openAddTransferLineChoice`, Manuel/Récurrent uniquement). `overlay-revmgr`/`overlay-revfilter` sont placés avant `overlay-picker` dans le shell (§12).

#### 6.7.7 Export / import
Comme tous les autres réglages (§7.2), `revenus` est un champ de plus dans le même payload `params` — `window.__SF_REVENUS.buildParamsPayloadPart()`/`applyParamsImportPart(revenus, mode)` branchés sur `buildParamsPayload()`/`applyParamsImport()`. Deux feuilles Excel dédiées (§7, §8) :

| Feuille | Colonnes | Alimente |
|---|---|---|
| **Revenus_Lignes** | `Personne · Catégorie · Ligne · Type · Catégorie récurrente · Signe` | lignes des deux personnes, dans l'ordre d'apparition |
| **Revenus_Montants** | `Personne · Catégorie · Ligne · Mois · Montant` | montants des lignes manuelles uniquement (récurrent et salaire synchronisé recalculés à l'import) |

**Deux onglets séparés, jamais fusionnés en un seul** : cardinalités différentes (une ligne par ligne définie vs une ligne par mois renseigné) — un onglet combiné aurait exigé soit des colonnes dupliquées, soit une lecture par position plutôt que par nom d'en-tête. Lues par nom, comme le reste de la feuille Paramètres.

**Réconciliation casse/accents silencieuse** (même principe que §7.1 bis) : "Catégorie récurrente" est alignée sur une catégorie déjà utilisée par une transaction existante (`U.alignToCanon`/`U.transactionFieldCanon`), pour qu'une faute de casse ne fasse pas échouer silencieusement le rattachement d'une ligne récurrente à ses transactions sources (`AGG.buildRecurringIndex`, §6.7.1). Par ailleurs, la recherche d'une "Ligne" dans `Revenus_Lignes` lors de la lecture de `Revenus_Montants` retombe sur une comparaison insensible casse/accents si la correspondance exacte échoue (`findRevLine`/`findRevTransferLine`) — sans cela, une même "Ligne" retapée avec une casse différente dans un seul des deux onglets perdrait silencieusement le montant du mois concerné plutôt que de créer un doublon.

`periodFrom`/`periodTo`/`groupFilter` ne voyagent que via la sauvegarde `.json` ou la synchro Gist (§7.2, §8), jamais via Excel (`groupFilter` conserve un garde-fou contre un tableau vide qui écraserait un filtre existant, cf. `applyParamsImportPart` — un tableau vide `[]` est "truthy" en JS).

**Colonne "Catégorie" conservée uniquement pour les sentinelles de transfert** — une ligne normale (destinée à `lines`) l'exporte vide ; seules les deux sentinelles réservées (`"(Transfert - vers)"`/`"(Transfert - depuis)"`) y ont un sens, pour router une ligne vers `transferOut`/`transferIn` à l'import.

**Fusion (mode "Ajouter")** : le bloc `revenus` complet (les deux personnes ensemble) suit la famille "adoption non destructive" (§7.2, famille 2) — adopté uniquement si rien n'est encore configuré localement pour aucune des deux personnes, jamais fusionné champ par champ. En mode "Remplacer" depuis Excel, gaté par `sheetsPresent.revenus` (§7.2) — un classeur ne contenant que, par exemple, `Budget` ne vide jamais une configuration Revenus existante. Une sauvegarde `.json`/synchro Gist applique toujours le remplacement si la clé `revenus` est présente.

---

## 7. Import

- Formats : `.xlsx`/`.xls`, `.csv` (données), `.json` (sauvegarde propre, v2). Colonnes de données (Date/Type/Montant) : formats acceptés en §1.1.
- `findDataSheetName()` détecte la feuille de données via ses en-têtes Date+Montant (ignore "Paramètres").
- **Six feuilles optionnelles**, indépendantes les unes des autres (aucune requise), lisibles depuis un classeur Excel de paramètres — toutes alimentent `params`, retourné par `parseParametresSheet(wb)` avec exactement la même forme que `buildParamsPayload()` (§7.2), donc consommées par le même `applyParamsImport(params, mode)` que la sauvegarde `.json`/la synchro Gist :

| Feuille | Colonnes | Alimente |
|---|---|---|
| **Paramètres** | `Groupe · Groupe de groupe · Type budget (Quotidien/Voyage) · Mois voyage · Date début (jours) · Date fin (jours) · [col. vide] · Compte · Type de compte · Acronyme · [col. vide] · Personne · Groupe de personnes · [col. vide] · Catégorie (classement forcé) · Type forcé (Revenus/Dépenses)` | `groupAliasRules`, `personAliasRules`, `accountTypes`, `accountAcronyms`, `groupKind`/`groupVoyageMonth`, `categoryTypeOverride`, `groupDateRange` |
| **Budget** | `Mois · Groupe · Catégorie · Montant · Remarque · Amorti (Oui/Non) · Vue (Perso/Couple/Les deux) · Vue uniquement (Oui/Non)` | `budget.adjustments` |
| **Budget_Salaires** | `Mois · Moi · Ma femme` | `budget.salaryEntries` |
| **Budget_Predefini** | `Groupe · Catégorie · Montant mensuel` | `budget.compare.predefined` |
| **Revenus_Lignes** / **Revenus_Montants** | voir §6.7.7 | `revenus.{moi,conjoint}` |

Trois colonnes vides intercalées dans **Paramètres** (après "Date fin (jours)", après "Acronyme", après "Groupe de personnes") — purement visuelles, ignorées à l'import, sans effet sur le mapping par nom des autres colonnes.

**Colonne "Mois voyage"** : vraie date Excel (toujours le 1ᵉʳ du mois, affichée `dd/mm/yyyy`, voir §8 pour l'écriture).

**Colonne "Groupe" de la feuille Paramètres — un seul bloc de colonnes contiguës** : `Groupe`, `Groupe de groupe`, `Type budget`, `Mois voyage`, `Date début/fin (jours)` se suivent directement. La colonne "Groupe" contient toujours le libellé complet exact (suffixe de date inclus) — sauf pour une règle générique volontairement créée sans occurrence précise (§2.1), auquel cas elle montre le nom tel quel. **Détection automatique** : si la valeur de "Groupe" contient un suffixe entre parenthèses (libellé complet), le nom de base est dérivé automatiquement (`U.baseGroupe`) et la règle est scopée à cette occurrence précise (`rule.full = valeur brute`). Type budget/Mois voyage/Bornes (jours) sont toujours indexés directement sur la valeur brute de "Groupe" (jamais dérivés), cohérent avec §2.1.

**Lecture des dates Excel** : `parseParametresSheet` lit "Mois voyage"/"Date début (jours)"/"Date fin (jours)" avec `raw:true` — une cellule de type date est renvoyée comme un vrai objet `Date`, sans ambiguïté jour/mois liée au format d'affichage de la cellule (`cellDates:true` étant actif à la lecture du classeur). Même principe pour la colonne `Date` de la feuille de transactions (§1.1) et pour "Date du soin"/dates de remboursement de Santé (§6.6).

**Ce principe ne s'applique jamais à un import CSV** (`raw` reste `false` pour ces mêmes lectures) : un CSV n'a aucune notion de format de cellule, et le texte brut tel que tapé (`JJ/MM/AAAA`, relu par `U.parseDateValue` avec la convention jour/mois de l'app, §1.1) y reste la lecture la plus fiable.

**Compte et Compte cible partagent un même espace de noms** : un compte existant côté source vaut aussi côté cible et inversement ; une correspondance/un remplacement choisi s'applique aux deux champs à la fois.

**Santé (§6.6) a son propre flux d'import Excel/CSV, entièrement séparé** — colonnes, dédoublonnage et écran dédiés ; réconciliation casse/accents silencieuse propre (§7.1 bis), sans écran de confirmation comme en §7.1. Les soins ne transitent par ce §7 que de façon transparente lors d'une restauration `.json` complète (§6.6, §8).

- Choix à l'import : Transactions+Paramètres / Transactions seules / Paramètres seuls (comme à l'export, §8) ; l'onglet Santé, s'il est présent dans le classeur, est détecté et importé **indépendamment de ce choix** (même principe qu'une sauvegarde `.json`, §6.6/§8). Puis Ajouter/Remplacer pour chaque volet (Santé partage le mode choisi pour les Transactions). En mode Ajouter, `Budget_Salaires` fusionne mois par mois ; en mode Remplacer, tous les mois importés remplacent intégralement les saisies locales. `Budget_Predefini` suit la même logique.
- `Store.normalizeSettings()` est systématiquement rappelée en fin d'import de paramètres (fichier ou synchronisation), pour migrer silencieusement tout réglage devenu obsolète.

### 7.1 Réconciliation (casse/accents + nouveaux éléments)

**Uniquement pour un import Excel/CSV avec des transactions déjà présentes** (jamais pour une restauration de sauvegarde `.json`, ni pour un premier import sur données vides). Avant l'écran de résumé habituel, `DATA.buildReconciliation(txns)` compare chaque valeur importée (**Catégorie, Personne, Compte + Compte cible, Groupe** — pas Remarque, pas Type) aux valeurs déjà présentes dans `Store.transactions`, comparaison insensible casse/accents (`U.normalize`) :

- **Correspondance trouvée** (ex. "restauration" importé, "Restauration" existant) → alignement automatique et silencieux sur la casse existante, mutation directe de `txns` en place. Un message informatif ("N valeur(s) alignée(s)") apparaît sur l'écran de résumé.
- **Aucune correspondance** → valeur véritablement nouvelle. Plusieurs variantes de casse d'une même valeur nouvelle dans le même import sont d'abord fusionnées entre elles (casse de la première occurrence fait foi), puis présentées comme une seule entrée nouvelle.
- S'il existe au moins une valeur véritablement nouvelle, un écran dédié (`renderImportReconcile`) les liste, groupées par champ, chacune avec le nombre d'occurrences et deux choix : **"Nouveau"** (par défaut) ou **"Remplacer par…"** (picker plein écran habituel, §6.2, pour la faire correspondre à une valeur déjà existante). "Continuer" applique les choix avant l'écran de résumé.

### 7.1 bis Réconciliation silencieuse des feuilles de réglages (Paramètres, Budget, Budget_Predefini, Revenus, Santé)

**Mécanisme distinct de celui du §7.1** (pas d'écran de confirmation, jamais de choix "Nouveau"/"Remplacer par…") : appliqué directement, en silence, à toute colonne de ces feuilles censée correspondre à une valeur déjà connue de l'application — pour éviter qu'une simple faute de casse/d'accent crée un réglage ou une ligne fantôme (nouvelle entrée) au lieu de mettre à jour l'existant. Sans correspondance, la valeur tapée est gardée telle quelle (comportement inchangé pour une valeur réellement nouvelle). Trois fonctions partagées portent ce mécanisme (`window.__SF_UTILS`) :
- `U.alignToCanon(valeur, table)` : aligne une valeur sur son équivalent existant si `U.normalize` les identifie, sinon la renvoie telle quelle.
- `U.transactionFieldCanon(champs)` : construit la table à partir d'un ou plusieurs champs déjà présents dans `Store.transactions` (ex. `["categorie"]`, ou `["compte","compteCible"]` pour regrouper compte source et cible dans le même espace de noms).
- `U.valueListCanon(valeurs)` : même principe à partir d'une simple liste de chaînes déjà connues (groupes budget déjà configurés, ou une des listes personnalisables de Santé).

Appliqué (`parseParametresSheet`, sauf Santé qui a son propre point d'entrée `rowToSanteItem`, §6.6) :
- **Paramètres** : Compte (Type de compte/Acronyme) sur un compte existant (source ou cible) ; Personne (Groupe de personnes) sur une personne existante ; Catégorie (classement forcé) sur une catégorie existante ; Groupe (Type budget/Mois voyage/Bornes/règles de groupe) sur le libellé littéral exact d'un groupe déjà présent dans les transactions — jamais sur un nom de base, pour respecter le principe du §2.1 (occurrence exacte).
- **Budget** (ajustements manuels) : Groupe aligné sur un groupe budget déjà configuré (`cfg.mainGroup`/`extraGroups`) ou, à défaut, sur un groupe littéral existant (`adjustmentGroupMatches` sait ensuite le résoudre via la hiérarchie) ; Catégorie alignée sur une catégorie déjà connue de ce groupe (`window.__SF_BUDGET.categoriesForGroup`, même liste que celle proposée dans la fiche "Budget prédéfini").
- **Budget_Predefini** : Groupe aligné strictement sur `cfg.mainGroup`/`extraGroups` (seule valeur permettant de retrouver la ligne au rendu, §6.3.5, depuis la suppression de la colonne "Clé") ; Catégorie alignée sur une catégorie déjà connue de ce groupe, même source que ci-dessus.
- **Revenus** (§6.7.7) : "Catégorie récurrente" alignée sur une catégorie existante ; repli casse/accents entre `Revenus_Lignes` et `Revenus_Montants` pour retrouver une "Ligne" (détail §6.7.7).
- **Santé** (§6.6) : Groupe aligné sur un groupe littéral existant ; Bénéficiaire/Devise/Médecin/Type de prestation/Statut/Mutuelle alignés sur une valeur déjà présente dans la liste personnalisable correspondante — appliqué avant le calcul de l'id (`buildSanteId`), pour qu'un réimport avec une casse différente ne crée pas de doublon du soin.

**Volontairement hors périmètre** : les catégories amorties (`amortCategories`) sont choisies via un sélecteur dans l'appli, jamais tapées à la main — pas exposées à ce risque.

### 7.2 Ce qui voyage réellement dans "params" (`buildParamsPayload()` / `applyParamsImport()`)

**Un seul et même payload `params`**, quelle que soit la porte d'entrée — sauvegarde `.json` (§8), synchronisation Gist (§10), ou les feuilles Excel (§7, `parseParametresSheet` construit un objet de forme identique). Toute correction ici s'applique automatiquement aux trois mécanismes à la fois. L'Excel (six feuilles, §7) ne peuple qu'un sous-ensemble de ces champs ; les champs qu'aucune feuille Excel ne porte sont simplement absents/`null`/vides dans l'objet retourné par `parseParametresSheet`, jamais activement écrasés (garde-fous détaillés ci-dessous).

**Inclus** : `groupAliasRules`, `personAliasRules`, `accountTypes`, `accountAcronyms`, `accountTypeOrder` (§6.1), `customBalances` (§6.1), `decimals` (§5), `groupKind`/`groupVoyageMonth`/`groupDateRange`, `categoryTypeOverride`, `budget.{adjustments,mode,scope,startMonth,endMonth,mainGroup,extraGroups,amortCategories,excludedAccounts,periodFrom,periodTo,salaryEntries,compare,mergeVoyages,mergeVoyagesGroups}`, `evolution.{view,monthBasis,merges}` (§6.4.5), `revenus.*` (§6.7.7 — bloc à part, sémantique de fusion détaillée là-bas).

**Ne voyage que par JSON/Gist, jamais par Excel** — pas de feuille Excel dédiée pour ces champs :
- `budget.{mode,scope,startMonth,endMonth,mainGroup,extraGroups,amortCategories,excludedAccounts,periodFrom,periodTo}`
- `budget.compare.{enabled,mode,periodFrom,periodTo}` (`budget.compare.predefined` reste porté par la feuille `Budget_Predefini`)
- `accountTypeOrder` (§6.1), `decimals` (§5)
- `customBalances` (§6.1)
- `evolution.{view,monthBasis,merges}` (§6.4.5)
- `revenus.{periodFrom,periodTo,groupFilter}` (§6.7.5, §6.7.7 — `revenus.{moi,conjoint}`, les lignes/montants eux-mêmes, restent portés par `Revenus_Lignes`/`Revenus_Montants`)

**Garde-fous anti-perte de données** (l'absence d'un champ dans l'Excel ne doit jamais l'écraser à vide/`null` en réimportant) :
1. `applyParamsImport`, mode "Remplacer" : `budget.periodFrom`/`periodTo` reprennent un repli sur la valeur déjà en place si l'import ne les fournit pas (`budgetSettings.periodFrom || Store.settings.budget.periodFrom`), même logique que `mode`/`scope`.
2. `window.__SF_REVENUS.applyParamsImportPart` : `revenus.groupFilter` n'est adopté que si le tableau importé est **non vide** — un tableau vide `[]` est "truthy" en JavaScript, un test de simple présence viderait le filtre à chaque import Excel.
3. `applyParamsImport`, mode "Remplacer" : chaque bloc de réglages Excel n'est écrasé que si `!params.sheetsPresent || sheetsPresent.xxx` — vrai systématiquement pour un payload JSON/Gist (`sheetsPresent` y est toujours absent), donc un remplacement JSON/Gist applique bien tous les champs sans exception ; pour un import Excel ciblé par onglet, seul le bloc dont la feuille est effectivement présente dans le classeur est écrasé.

**Volontairement exclus** (état d'affichage/session courant, jamais synchronisé — même principe que les filtres Transactions, `Store.settings.filters`) : `theme`, sélection de filtres courante (Transactions comme Santé), `txnGroupsExpanded`/`recurrent.groupsExpanded`/`santeGroupsExpanded`, `budget.{detailView,monthlyView,displayMode}`, `recurrent.{periodFrom,periodTo}`, `evolution.cumulative` (§6.4.2 bis), `hideZeroAccounts` (§6.1). `Store.settings.sync` (jeton, ID de Gist) est exclu pour une raison différente et non négociable : des identifiants ne doivent jamais se retrouver dans un fichier de sauvegarde ou repartir vers le Gist qu'ils servent à authentifier.

**Sémantique de fusion (mode "Ajouter", tous mécanismes confondus)** — trois familles de comportement, jamais mélangées pour un même champ :
1. **Fusion par dédoublonnage** (listes de règles/entités identifiables individuellement) : `groupAliasRules`, `personAliasRules`, `budget.adjustments`, `evolution.merges.*` — chaque élément importé est ajouté sauf s'il existe déjà localement (`mergeRuleList`, clé de signature propre à chaque type de liste).
2. **Adoption non destructive** (réglage global unique, jamais fusionné champ à champ) : `accountTypeOrder`, `customBalances`, `budget.excludedAccounts`, la plupart des champs de `budgetSettings` — la valeur importée n'est adoptée que si rien n'est encore configuré localement ; elle n'écrase jamais un réglage déjà présent.
3. **Jamais touché par un import "Ajouter"** : `evolution.view`/`evolution.monthBasis` et `decimals` en mode fusion — seul un import "Remplacer" les modifie.

**Mode "Remplacer" (`mode === "replace"`) : deux comportements distincts selon la porte d'entrée**, tous deux dans `applyParamsImport` :
- **Sauvegarde `.json` / synchro Gist** (`payload.params` n'a jamais de champ `sheetsPresent`) : écrase toujours tout sans distinction, quel que soit le champ.
- **Import Excel ciblé par onglet** (`params.sheetsPresent` présent) : chaque bloc de réglages n'est écrasé que si la feuille Excel correspondante est effectivement présente dans le classeur (`sheetsPresent.parametres`, `.budget`, `.revenus`) — un classeur ne contenant que, par exemple, `Budget` ne vide jamais `groupAliasRules`/`accountTypes`/etc. Les champs qui ne voyagent par aucune feuille Excel (liste ci-dessus) ne sont jamais écrasés par un import Excel.

---

## 8. Export
Menu → Excel (.xlsx) ou sauvegarde (.json).

**Export Excel — trois cases à cocher indépendantes (Transactions / Paramètres / Santé), 1 à 3 sélectionnables à la fois** (`openExportExcelScopeChoice`) — la case Santé n'apparaît que si `Store.santeItems` n'est pas vide. Un seul classeur est produit, avec une feuille par volet coché. Nom de fichier : suffixe reflétant la sélection (`_transactions`, `_parametres`, `_sante`, ou combinaisons), aucun suffixe si les trois sont cochées. Export JSON (sauvegarde `.json`) : choix exclusif à trois options Transactions+Paramètres / Transactions seules / Paramètres seuls — Santé voyage toujours avec le JSON quel que soit ce choix (§6.6).

**Dates au format Excel natif JJ/MM/AAAA** : toutes les colonnes à précision jour sont écrites comme de vraies cellules de type date (`XLSX.utils.aoa_to_sheet(data, {cellDates:true})`), format d'affichage forcé à `dd/mm/yyyy` (`setColumnDateFormatDDMMYYYY`). **"Mois voyage" (§2.4)** : vraie cellule date au 1ᵉʳ du mois (`voyageMonthToDateObj`, ex. mois interne `"2024-01"` → cellule affichée `01/01/2024`) ; le stockage interne (`Store.settings.groupVoyageMonth`) reste inchangé, seule la représentation Excel change. À l'import, `U.parseDateValue` lit indifféremment une vraie cellule date, un numéro de série Excel, ou du texte JJ/MM/AAAA(-AA), et n'en retient que l'année/le mois (le jour est ignoré).

**Colonne Type en français** : la feuille de transactions exporte `Revenu`/`Dépense`/`Transfert` (au lieu de l'anglais interne `Income`/`Expense`/`Transfer`), cohérent avec le format d'import (§1.1). Les montants restent des cellules numériques signées classiques.

**Feuille `Budget_Predefini`** : colonnes `Groupe`/`Catégorie`/`Montant mensuel` uniquement (§6.3.5) — plus de colonne `Clé` ni `Vue`, la clé interne est reconstruite à l'import depuis Groupe+Catégorie, avec réconciliation casse/accents silencieuse (§7.1 bis).

**Santé (§6.6)** garde son export Excel dédié (menu propre à l'onglet, colonnes identiques) et peut aussi être incluse dans l'export Excel principal via la case à cocher (même fonction `buildSanteExportRows()` partagée). Toujours incluse dans la sauvegarde `.json` complète, quel que soit le périmètre Transactions/Paramètres choisi pour le JSON.

---

## 9. Stockage local
IndexedDB (`sf_finance_db`) pour les transactions, repli `localStorage` si indisponible. `Store.settings` (thème, règles, types de compte, acronymes, réglages budget/évolution/récurrent, `groupDateRange`, `sync`, `sante`/`santeFilters`…) en `localStorage` (`sf_settings_v1`).

Les soins de l'onglet Santé (§6.6) suivent le même principe dans une clé IndexedDB séparée (`sante_items`, repli `localStorage` `sf_sante_items_v1`), avec sa propre gestion d'erreur (`Store.lastSanteSaveError`, même mécanique que pour les transactions).

---

## 10. Synchronisation manuelle (GitHub Gist)

Mécanisme de transfert entre appareils alternatif à l'export/import de fichier, sans backend : sauvegarde/récupération du payload complet (transactions + paramètres + soins et réglages Santé, §6.6 — même format que la sauvegarde `.json` §7/§8) sur un Gist GitHub privé, via l'API `api.github.com/gists/{id}` appelée directement depuis le navigateur. Les soins Santé suivent la même règle "distant/local/hash" que les transactions ci-dessous, sans fusion possible à la récupération (remplacement complet).

**Configuration** (`Store.settings.sync = {token, gistId, filename, lastSyncedAt, lastSyncedHash}`) : jeton d'accès personnel (scope **`gist`** uniquement recommandé). L'**ID du Gist est optionnel** : laissé vide, le premier "Synchroniser" crée automatiquement un nouveau Gist secret et mémorise l'ID retourné.

**Synchroniser** (`smartSync`) : sens **automatique** dans les deux directions.
1. Calcule trois hash : local actuel, distant actuel, `lastSyncedHash`.
2. Seul le distant a changé → récupère automatiquement. Seul le local a changé → sauvegarde automatiquement. Identiques → rien à faire.
3. **Les deux ont changé indépendamment** → modale de conflit demandant explicitement "Garder cet appareil" ou "Utiliser le Gist".

**Forcer sauvegarde** / **Forcer récupération** : action à sens unique, pour un rattrapage manuel.

**Indicateur** (`#btn-sync`) : icône + couleur selon l'état (`computeStatus()`). Module `window.__SF_SYNC` : `{open, updateIndicator, pushToGist, pullFromGist, smartSync}`.

---

## 11. Interface (style One UI)
Palette bleue, cartes arrondies, mode sombre auto/forçable. Numéro de version affiché en petit sous le titre "Suivi financier" (`APP_VERSION` sans le préfixe `v `, `buildShell`). Bouton flottant unique à 4 icônes (Affichage/Recherche/Filtre/Plus, §6.2) visible sur les onglets Transactions et Santé (§6.6) uniquement — absent de Revenus (§6.7), qui n'en a pas besoin. Deux onglets secondaires (Revenus §6.7, Santé §6.6) accessibles via un bouton "⋯" en fin de barre de navigation — pattern pensé pour accueillir d'éventuels futurs onglets secondaires sans réencombrer la barre principale (§6.6, §6.7, §12). Génération de données d'exemple depuis les écrans vides. Suppression protégée par confirmation (transaction, soin Santé, ajustement budgétaire, règle groupe/personne, catégorie/ligne de Revenus — `APP.confirmAction`, §6.2 ; supporte du texte multi-lignes via `\n`). Réinitialisation complète dans le Menu (confirmation requise, efface aussi les soins Santé).

Classe CSS `.chev-rot` : fait pivoter une flèche `<svg>` de 180° quand son `<details>` parent est ouvert. Pilules de sélection multiple (`.pill-toggle`/`.pill-toggle-wrap`) : utilisées dans l'onglet Budget pour les comptes exclus et les catégories amorties, et dans l'onglet Santé pour le statut Terminé et la sélection rapide de la part personnelle (0/50/100 %, §6.6). Classe `.txn-tap-row` (fiche transaction, §6.2, réutilisée par la fiche soin de Santé, §6.6) : ligne tactile générique (fond au tap via `:active`) — réutilisable pour toute future ligne "détail cliquable pour éditer".

**Colonnes figées (tableau Évolution, §6.4.3, et Récurrent, §6.5)** : pattern réutilisable si un futur tableau a besoin du même comportement — `<table>` HTML classique (pas de grille CSS), `box-sizing:border-box` obligatoire sur toutes les cellules, largeurs `width`/`min-width`/`max-width` identiques sur chaque colonne figée, et `left` de la 2ᵉ colonne figée = largeur exacte de la 1ʳᵉ (Évolution fige Libellé + Total ; Récurrent ne fige que Libellé). Défilement horizontal au-delà des colonnes figées : `U.enableDragScroll` (glisser-déposer souris + tactile natif, §6.4.3/§12), partagé par Évolution, Récurrent et la vue mensuelle du Budget. Un en-tête sticky "flottant" au-dessus d'une carte (§6.3.1) doit être un enfant direct du conteneur qui englobe tout le contenu à couvrir par le sticky, pas seulement de la carte elle-même.

**Fiches de réglages du Budget qui se re-rendent en place** (Amortissement, Comparaison, Comptes exclus, Soldes personnalisés, Budget prédéfini) : chaque `render()` local doit mémoriser puis restaurer la position de défilement autour de la réécriture de `modal.innerHTML` (`const prevScroll = modal.scrollTop; ...; modal.scrollTop = prevScroll;`) — sinon toute saisie ou clic dans la fiche la ramène en haut. Si la fiche a son propre conteneur défilant interne (`overflow-y:auto` distinct de `modal-detail`), lui donner un id et restaurer aussi son `scrollTop` séparément (cf. `#pbf-list` dans le budget prédéfini). Tout nouveau sous-écran du Budget suivant ce pattern doit reprendre cette convention.

---

## 12. Architecture technique

| Module (IIFE, ordre d'exécution) | Expose | Rôle |
|---|---|---|
| Utils | `window.__SF_UTILS` | formatage, dates, icônes SVG, `baseGroupe()`, `baseRemarque()` (nom de transaction sans suffixe de date, §6.5), `computeTransform()`, `enableDragScroll(el)` (glisser-déposer horizontal, §11), `decimalsToggleBtnHtml()`/`toggleMoneyDecimals()` (§5), `hideZeroAccountsBtnHtml()`/`toggleHideZeroAccounts()`/`isZeroBalance(bal)` (§6.1) |
| Store + Import/Parsing | `window.__SF_STORE`, `window.__SF_DATA` | IndexedDB/localStorage, parsing fichiers, `parseParametresSheet(wb)` (lit les feuilles Excel de paramètres, §7.2, même forme de sortie que `buildParamsPayload()`), dédup, démo, `normalizeSettings()` |
| Filtres + Agrégations | `window.__SF_FILTERS`, `window.__SF_AGG` | état des filtres (`getActiveChips`/`removeChip` exposés sur `Filters`, §4), `getGroupeDisplay`/`getPersonneDisplay`/`getAccountTypeDisplay`/`getAccountAcronym`, `getGroupKind`/`getGroupVoyageMonth`/`getGroupDateRange`/`getGroupDayStats`/`getCalcMonthYM` (indexés sur le libellé complet, §2.1), `groupeHierarchyChrono`/`getGroupStartDate`/`groupStartDatesMap`/`groupLabelSuffix`/`parseGroupSuffixDate` (tri chronologique, §2.1 bis), `GROUP_DATE_TODAY_SENTINEL`/`resolveGroupDateToValue` (§2.5), `labelResolvesToSuperGroupe` (résolution groupe de groupe d'un libellé, utilisée par le matching des ajustements du Budget, §6.3.6), `distinctMonths(txns, "indiquee"|"transformee")` (§4), calculs de soldes/totaux/séries, **moteur Évolution** (`buildEvolutionIndex`, `itemTotalValue`/`sectionTotalValue`, `itemYearValue`/`sectionYearValue`, `itemMonthValue`/`sectionMonthValue`, `evolutionYearsList`, §6.4 ; `itemCumulativeYearValue`/`sectionCumulativeYearValue`, `itemCumulativeMonthValue`/`sectionCumulativeMonthValue` pour le mode Cumulé, §6.4.2 bis), **moteur Récurrent** (`buildRecurringIndex(periodFrom, periodTo, groupFilter?)`, `recurringItemTotal`/`recurringSectionTotal`, `recurringItemMonthValue`/`recurringSectionMonthValue`, `recurringItemYearLumpValue`/`recurringSectionYearLumpValue`, `recurringColumns`, §6.5 — `groupFilter` optionnel, §6.7.5, teste `t.groupeBase` **et** `getGroupeDisplay(t)`) |
| Contrôleur | `window.__SF_APP` | shell, navigation (Comptes/Transactions/Budget/Évolution/Récurrent dans la barre principale + Revenus/Santé derrière "⋯", `currentTab` par défaut `"accounts"`, `MORE_TABS`/`MORE_TAB_KEYS`/`openMoreMenu()`) ; thème, état partagé (dont `txnNavActive`, §6.2), `confirmAction(opts)` (§6.2) ; `closeOverlay("overlay-detail")` réinitialise `txnNavActive` et les gestionnaires de balayage tactile de `modal-detail` (§6.2) |
| Vue Comptes/Transactions | `window.__SF_VIEWS.*` | Comptes (regroupement par type, réorganisable, copier pour Excel) ; Transactions (§6.2) : `renderTransactions`, `computeTxnGroups`/`renderTxnGroupCards` (rendu unifié des 11 modes), `flattenTxnGroupIds(arr)` (base de la navigation précédent/suivant), `openDisplayMenu`, `openTxnDetail(id, navIds?)`/`openTxnNew` (→ `renderTxnDetail(id, seed, navList)` commun), `openInlineEditor` (§6.2), `categoriePickerHierarchy()`, `setupTxnInfiniteScroll` |
| Vue Budget | idem, dans `window.__SF_VIEWS.renderBudget` | 2 modes (Amort.+ajouts / Ajouts), vue détail 3 colonnes, `computeGroupBudget()`, `groupSelectionMatches()`/`adjustmentGroupMatches()` (correspondance groupe/ajustement avec repli hiérarchique, §2.1/§6.3.6), `salaryMonthBreakdown()`, `groupPickerHierarchy()` (picker groupe principal/additionnel, alpha, §6.3.2), `adjustmentGroupHierarchy()` (picker groupe d'un ajustement, chronologique + icône type, §6.3.6) |
| Vue Évolution | idem, dans `window.__SF_VIEWS.renderEvolution` | 4 dimensions, drill-down année/mois, mode cumulé (§6.4.2 bis), fusions de sections (§6.4.5), `openEvolutionTxnSheet()` (§6.4) |
| Vue Récurrent | idem, dans `window.__SF_VIEWS.renderRecurrent` | tableau catégorie → transaction sans date, colonne année spéciale, filtre de période, `openRecurringTxnSheet()` (§6.5) |
| Vue Revenus | idem, dans `window.__SF_VIEWS.renderRevenus` ; `window.__SF_REVENUS` (`buildParamsPayloadPart`/`applyParamsImportPart`) | Moi/Ma femme, catégories + transferts miroir, `openRevManager()`/`openRevFilterSheet()`/`openRevYearEditor()` (§6.7) |
| Feuille de filtre | `window.__SF_FILTERSHEET` | menu de filtre en deux étapes ; `moisGroupedHtml`/`wireMoisGrouped` et `comptesGroupedHtml`/`wireComptesGrouped` (§4), `quickSelectRowHtml`/`wireQuickSelect`, `activeFiltersSectionHtml`, `NESTED_DIM_CONFIG` (filtre Groupe chronologique, §2.1 bis — celui de Revenus, §6.7.5, réimplémente le même rendu sans passer par `NESTED_DIM_CONFIG`) |
| Import/Menu | `window.__SF_IMPORT`, `window.__SF_MENU` | import, export, réinitialisation ; `buildParamsPayload()`/`applyParamsImport()` exposés pour réemploi (Synchronisation §10, `parseParametresSheet` §7.2 — un seul payload `params`, trois portes d'entrée) ; `DATA.buildReconciliation(txns)`/`renderImportReconcile()` (§7.1) |
| Synchronisation | `window.__SF_SYNC` | sauvegarde/récupération bidirectionnelle via GitHub Gist (§10) |
| Onglet Santé (IIFE unique) | `window.__SF_SANTE`, `window.__SF_SANTE_DATA`, `window.__SF_SANTE_FILTERS`, `window.__SF_SANTE_FILTERSHEET`, `window.__SF_SANTE_SETTINGS` | modèle, stockage, filtres, agrégations, vues et réglages de §6.6 ; `computeMaPart`/`buildSanteId`/`sanitizeItem` (`__SF_SANTE_DATA`), `rowToSanteItem`/`exportSanteExcel`/`buildSanteExportRows` (Excel dédié §6.6, aussi consommée par l'export principal §8), `importSanteItems` (fusion/remplacement) |
| `createAliasManager(cfg)` (fonction globale) | — | fabrique réutilisée par Groupes et Personnes ; vue hiérarchique unique, édition par occurrence, dates de jours (+ case "toujours aujourd'hui", §2.5), icônes de type, scroll préservé, `obsoleteRules()`/nettoyage des règles obsolètes (§2.4), suppression protégée par `APP.confirmAction`. Toutes les requêtes DOM internes (`am-*`) doivent être scopées au `sheet` local, jamais à `document` (§2.4) |
| Gérer les groupes | `window.__SF_GROUPMGR` | `hasDates:true, showKind:true` |
| Gérer les personnes | `window.__SF_PERSONMGR` | `hasDates:false` |
| Gérer les types de compte | `window.__SF_ACCTTYPE` | type + acronyme (§3) |

**Pour ajouter une dimension similaire à Groupe/Personne** : dupliquer le schéma `getXDisplay`/`findXAliasRule`/`distinctBaseX`/`xHierarchy` dans le module Filtres+Agrégations, appeler `createAliasManager({...})`, ajouter une entrée `NESTED_DIM_CONFIG` pour le filtre imbriqué.

**Pour ajouter une dimension à l'onglet Évolution** : ajouter un cas dans `buildEvolutionIndex` (section/item pour cette dimension) et dans `txnMatchesDimTarget` (filtrage de la feuille de détail) — les deux doivent rester cohérents entre eux.

**Pour un picker de fiche transaction groupé/trié différemment** (§6.2) : `renderPickerList` accepte `opts.hierarchy = {sections: Map, flat: [], preSorted?, selectableHeaders?, expandedByDefault?, kindFn?}` — sans `preSorted`, les sections et leurs enfants sont re-triés alphabétiquement par la fonction elle-même (cas du picker de groupe principal/additionnel du Budget, `groupPickerHierarchy`, §6.3.2) ; avec `preSorted:true`, l'ordre fourni par l'appelant est respecté tel quel (tri chronologique des groupes, §2.1 bis — fiche transaction, filtre Groupe, picker de groupe d'un ajustement §6.3.6 — et picker Catégorie qui impose l'ordre Revenus puis Dépenses). `selectableHeaders:false` désactive la sélection de l'en-tête de section (Catégorie, où l'en-tête est un type) ; `expandedByDefault:true` déplie toutes les sections dès l'ouverture ; `kindFn(base)` (optionnel, icône Voyage/Vie quotidienne du picker d'ajustement §6.3.6) retourne un HTML à préfixer devant le libellé de chaque groupe précis, jamais devant un en-tête de section.

**Pour ajouter un futur onglet secondaire derrière le bouton "⋯"** (§6.6, §11) : ajouter une entrée à `MORE_TABS`/`MORE_TAB_KEYS` (module Contrôleur), un nouveau `view-*` dans le shell, inclure le nouvel onglet dans `showFloat` si le bouton flottant s'applique, brancher `refreshCurrentTab()` et le dispatch des 4 boutons selon `currentTab` — l'onglet Santé sert de modèle complet, y compris pour une liste déroulante personnalisable avec ajout à la volée (`PICKER.openPicker({..., allowNew:true})` + persistance dans `Store.settings`) et un export/import Excel dédié (§6.6) qui reste distinct des écrans du flux principal (§7/§8) tout en partageant son mapping de colonnes (`buildSanteExportRows`) et sa fonction de lecture (`rowToSanteItem`).

---

## 13. Limites connues
- Pas de synchronisation automatique entre appareils : export/import JSON ou Excel, ou synchronisation manuelle GitHub Gist (§10), comme mécanismes de transfert. Résolution de conflit au hash global uniquement (pas de fusion champ par champ).
- Détection "Paramètres"/"Budget"/etc. et mapping de colonnes basés sur des noms d'en-tête normalisés, pas sur la position.
- Règles de groupe : résolution "occurrence précise d'abord, règle générique en repli" (§2.1), toujours manuelle (aucune déduction géographique/temporelle automatique). Pas de fenêtre de dates sur les règles elles-mêmes (§2.5) — seule la présence/absence d'un libellé complet distingue les occurrences.
- `groupDateRange` est une donnée déclarative simple (aucune vérification de cohérence avec les dates réelles des transactions du groupe, hormis le garde-fou jours ≤ 0). La date de fin dynamique "aujourd'hui" (§2.5) ne s'applique qu'à `dateTo` ; pas d'équivalent pour `dateFrom`.
- Onglet Évolution : les fusions de sections (§6.4.5) sont indépendantes par dimension et non partagées avec un quelconque mécanisme de groupe de groupe/personne — fusionner deux sections dans Évolution ne crée aucune règle `groupAliasRules`/`personAliasRules` et inversement. Une nouvelle fusion ne peut être créée que via JSON/Gist (§7.2) — l'onglet ne permet que d'en supprimer une existante.
- `evolution.cumulative` (§6.4.2 bis) et `hideZeroAccounts` (§6.1) sont des réglages locaux à l'appareil, absents de `buildParamsPayload` (§7.2) : un changement d'appareil réinitialise ces deux affichages à leur valeur par défaut. `evolution.view`/`monthBasis`/`merges` sont bien inclus dans `buildParamsPayload` et voyagent via `.json`/Gist, mais pas via l'Excel (§7.2).
- `AGG.getGroupStartDate` (§2.1 bis) interprète le suffixe entre parenthèses d'un groupe (`AA-MM-JJ`/`AA-MM_MM`/`AA-MM`/`AAAA`) uniquement en repli, quand aucune date de début n'est configurée manuellement — un suffixe dans un format non reconnu retombe silencieusement sur la date de la transaction la plus ancienne du groupe, sans avertissement.
- Réconciliation d'import interactive (§7.1) : ne couvre que Catégorie/Personne/Compte/Compte cible/Groupe de la feuille de transactions — pas Remarque, pas Type. Ne s'exécute que pour un import Excel/CSV avec des transactions déjà présentes ; une restauration `.json` ne passe jamais par cette étape. Les feuilles de réglages (Paramètres/Budget/Budget_Predefini/Revenus) et l'import Santé ont leur propre réconciliation silencieuse, sans écran de confirmation (§7.1 bis) — mécanisme distinct, jamais de choix "Nouveau"/"Remplacer par…".
- Onglet Santé (§6.6) : la fusion à l'import ignore silencieusement un soin dont l'id existe déjà — elle ne met jamais à jour un soin existant (ex. un `termine` basculé localement puis un import "Ajouter" du même fichier source ne l'écrase pas).
- Le filtre Groupe de Santé partage les comptages du filtre Groupe des Transactions (§2.1 bis, §6.6) — un groupe peut donc apparaître avec un nombre qui ne correspond à aucun soin si aucun n'y est encore rattaché.
- Santé : pas de taux de change — `montantDevise` est un champ informatif indépendant de `montantEur`, jamais calculé ou converti automatiquement.
- Revenus (§6.7) : une ligne récurrente est reliée par `(recurCategorie, recurLabel)` — une catégorie et un nom de transaction sans date, comme un item de l'onglet Récurrent (§6.5). Si la catégorie ou le texte de la Remarque change ensuite côté transactions, le lien cesse silencieusement de correspondre et la ligne affiche 0, sans avertissement. Le filtre Groupe de Revenus (§6.7.5) ne s'applique qu'aux lignes récurrentes ; aucun moyen de restreindre une ligne manuelle ou salaire synchronisé à un groupe.
- Détection automatique de la feuille "Santé" dans le flux d'import principal (§7) : le mode Ajouter/Remplacer appliqué aux soins suit toujours celui choisi pour les **transactions** (`chosenMode`), même si le volet Transactions n'est pas importé — pas de choix Ajouter/Remplacer dédié aux soins sur cet écran, contrairement à l'écran d'import Santé dédié (§6.6).
- Navigation précédent/suivant de la fiche transaction (§6.2) : la liste de navigation est figée à l'ouverture de la fiche et ne protège que le temps où celle-ci reste affichée — fermer la fiche puis rouvrir une transaction repart d'une liste recalculée sur le filtre à jour. Aucune navigation n'est disponible pour les fiches ouvertes depuis un autre onglet (Budget, Récurrent, Évolution, Santé).
- Ajustements manuels du Budget (§6.3.6) : la correspondance avec une carte remonte d'un groupe précis vers son groupe de groupe (§6.3.6), mais jamais dans l'autre sens — un ajustement posé directement sur un groupe de groupe ne redescend pas automatiquement sur une carte configurée sur l'un de ses groupes enfants précis, pour éviter toute ambiguïté si plusieurs cartes enfants existent.
- Testé via Playwright (Chromium headless) quand l'environnement le permet, sinon relecture manuelle systématique + `node --check` sur le script extrait ; SheetJS nécessite un accès internet au premier chargement.
