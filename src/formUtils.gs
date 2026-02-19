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
 * 旧形式の工番関連アイテムを全て削除し、
 * シンプルなテキスト欄「工番1」「工番2」「工番3」に差し替える。
 * WEBアプリのモーダルで選択した工番コードがそのままプリフィルされる。
 * 全て任意（間接業務の方は工番不要）。
 *
 * @param {Form} form - 対象フォーム
 */
function ensureOrderItems_(form) {
  // ---- 旧形式のアイテムを削除 ----
  const oldTitles = [
    Q.ORDER,                    // 旧「工番」プルダウン
    Q._OLD_ORDER_PREFIX_1, Q._OLD_ORDER_NUMBER_1,
    Q._OLD_ORDER_PREFIX_2, Q._OLD_ORDER_NUMBER_2,
    Q._OLD_ORDER_PREFIX_3, Q._OLD_ORDER_NUMBER_3,
  ];

  // 「追加の工番を入力しますか？」「さらに追加の〜」等の条件分岐アイテムも削除
  const additionalPatterns = [
    /追加.*工番/,
    /さらに.*工番/,
  ];

  let insertIndex = -1;

  // 逆順で削除（インデックスずれ防止）
  const items = form.getItems();
  for (let i = items.length - 1; i >= 0; i--) {
    const title = normalize_(items[i].getTitle());
    const isOld = oldTitles.some(t => t && title === t);
    const isAdditional = additionalPatterns.some(p => p.test(title));
    if (isOld || isAdditional) {
      if (insertIndex < 0 || items[i].getIndex() < insertIndex) {
        insertIndex = items[i].getIndex();
      }
      form.deleteItem(items[i]);
    }
  }

  // ---- 新形式のテキスト欄を作成（既存なら何もしない） ----
  const newQs = [Q.ORDER_1, Q.ORDER_2, Q.ORDER_3];
  const newItems = [];

  for (const qTitle of newQs) {
    const existing = findItemByTitleOrNull_(form, qTitle);
    if (existing) continue; // 既にある

    const item = form.addTextItem();
    item.setTitle(qTitle);
    item.setRequired(false);
    item.setHelpText('WEBアプリのモーダルから選択するか、工番コードを直接入力してください（任意）');
    newItems.push(item);
  }

  // 挿入位置が分かっていれば移動
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
