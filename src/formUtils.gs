// ====== フォームの質問アイテム取得（タイトルで検索） ======

function findItemByTitle_(form, title) {
  const items = form.getItems();
  for (const it of items) {
    if (normalize_(it.getTitle()) === title) return it;
  }
  throw new Error(`フォーム内に質問が見つかりません: "${title}"（テンプレに同名の質問を作ってください）`);
}

/**
 * タイトルで質問を探す（見つからなければ null を返す版）
 */
function findItemByTitleOrNull_(form, title) {
  const items = form.getItems();
  for (const it of items) {
    if (normalize_(it.getTitle()) === title) return it;
  }
  return null;
}

/**
 * 旧「工番」プルダウンを削除し、プレフィックス選択＋5桁番号入力（×3件）に差し替える。
 * 既に新形式の質問が存在する場合はプレフィックス選択肢の更新のみ行う。
 *
 * @param {Form} form - 対象フォーム
 * @param {string[]} prefixes - プレフィックス一覧（例: ['LW','LWTS','LWZ']）
 */
function ensureOrderPrefixItems_(form, prefixes) {
  if (!prefixes || prefixes.length === 0) {
    Logger.log('WARN: プレフィックス候補が空のためスキップ');
    return;
  }

  // 旧「工番」プルダウンが残っていれば削除（挿入位置を記憶）
  let insertIndex = -1;
  const oldItem = findItemByTitleOrNull_(form, Q.ORDER);
  if (oldItem) {
    insertIndex = oldItem.getIndex();
    form.deleteItem(oldItem);
  }

  // 3件分の定義（1件目は必須、2,3件目は任意）
  const pairs = [
    { prefix: Q.ORDER_PREFIX_1, number: Q.ORDER_NUMBER_1, required: true },
    { prefix: Q.ORDER_PREFIX_2, number: Q.ORDER_NUMBER_2, required: false },
    { prefix: Q.ORDER_PREFIX_3, number: Q.ORDER_NUMBER_3, required: false },
  ];

  // 5桁数字バリデーション（全角/半角対応）
  const digitValidation = FormApp.createTextValidation()
    .setHelpText('5桁の数字を入力してください（例：01234）')
    .requireTextMatchesPattern('^[0-9\uFF10-\uFF19]{5}$')
    .build();

  const newItems = [];

  for (const pair of pairs) {
    // --- プレフィックス（プルダウン） ---
    const existingPrefix = findItemByTitleOrNull_(form, pair.prefix);
    if (existingPrefix) {
      // 既に存在 → 選択肢のみ更新
      setDropdownChoices_(existingPrefix, prefixes);
    } else {
      const prefixItem = form.addListItem();
      prefixItem.setTitle(pair.prefix);
      prefixItem.setRequired(pair.required);
      prefixItem.setChoiceValues(prefixes);
      newItems.push(prefixItem);
    }

    // --- 番号（短文テキスト） ---
    const existingNumber = findItemByTitleOrNull_(form, pair.number);
    if (!existingNumber) {
      const numberItem = form.addTextItem();
      numberItem.setTitle(pair.number);
      numberItem.setRequired(pair.required);
      numberItem.setValidation(digitValidation);
      newItems.push(numberItem);
    }
  }

  // 新規追加したアイテムを旧「工番」があった位置に移動
  if (insertIndex >= 0 && newItems.length > 0) {
    for (let i = 0; i < newItems.length; i++) {
      form.moveItem(newItems[i], insertIndex + i);
    }
  }
}

function setDropdownChoices_(item, choices) {
  // DropdownItem or ListItem対応（テンプレの形式に合わせる）
  const type = item.getType();
  if (type === FormApp.ItemType.LIST) {
    const li = item.asListItem();
    li.setChoiceValues(choices);
    return;
  }
  if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
    const mi = item.asMultipleChoiceItem();
    mi.setChoiceValues(choices);
    return;
  }
  if (type === FormApp.ItemType.DROP_DOWN) {
    // ※DROP_DOWN は ItemType に存在しない場合があるため保険（環境差）
    item.asListItem().setChoiceValues(choices);
    return;
  }
  throw new Error(`この質問はプルダウン/ラジオではありません: "${item.getTitle()}" type=${type}`);
}
