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
  // ---- 旧形式の工番アイテムを検出して削除 ----
  const oldExactTitles = new Set([
    Q.ORDER,                    // 旧「工番」プルダウン
    Q._OLD_ORDER_PREFIX_1, Q._OLD_ORDER_NUMBER_1,
    Q._OLD_ORDER_PREFIX_2, Q._OLD_ORDER_NUMBER_2,
    Q._OLD_ORDER_PREFIX_3, Q._OLD_ORDER_NUMBER_3,
  ].filter(Boolean));

  // パターンマッチで削除する項目
  // - 「追加の工番を入力しますか？」「さらに追加の〜」等の条件分岐
  // - 条件分岐のセクション見出し（PAGE_BREAK / SECTION_HEADER）で
  //   タイトルが「工番」を含むもの
  const deletePatterns = [
    /追加.*工番/,
    /さらに.*工番/,
  ];

  // 新形式のタイトル（TEXT型でなければ削除対象）
  const newTitleSet = new Set([Q.ORDER_1, Q.ORDER_2, Q.ORDER_3]);

  let insertIndex = -1;
  let needsMigration = false;

  // まず旧項目が存在するか確認（存在しなければスキップ→高速化）
  const items = form.getItems();
  for (const it of items) {
    const title = normalize_(it.getTitle());
    if (oldExactTitles.has(title) || deletePatterns.some(p => p.test(title))) {
      needsMigration = true;
      break;
    }
    // 「工番1」等のタイトルがTEXT以外（セクション見出し等）→ 要マイグレーション
    if (newTitleSet.has(title) && it.getType() !== FormApp.ItemType.TEXT) {
      needsMigration = true;
      break;
    }
  }

  if (!needsMigration) {
    // 新形式のテキスト欄が既に3つ揃っているか確認
    const existing = [Q.ORDER_1, Q.ORDER_2, Q.ORDER_3].filter(
      q => findItemByTitleOrNull_(form, q) !== null
    );
    if (existing.length === 3) return; // 完全に移行済み → 何もしない
  }

  // ---- 旧項目を逆順で削除（インデックスずれ防止） ----
  const allItems = form.getItems();
  for (let i = allItems.length - 1; i >= 0; i--) {
    const title = normalize_(allItems[i].getTitle());
    const itemType = allItems[i].getType();

    // 完全一致で旧項目
    const isOldExact = oldExactTitles.has(title);

    // パターンマッチ（追加の工番〜、さらに〜）
    const isPattern = deletePatterns.some(p => p.test(title));

    // セクション見出し/ページ区切りで工番関連のもの
    const isSectionHeader = (
      (itemType === FormApp.ItemType.PAGE_BREAK || itemType === FormApp.ItemType.SECTION_HEADER) &&
      /工番/.test(title)
    );

    // 新タイトルだがTEXT以外（セクション見出し等が残っている場合）
    const isWrongType = newTitleSet.has(title) && itemType !== FormApp.ItemType.TEXT;

    if (isOldExact || isPattern || isSectionHeader || isWrongType) {
      if (insertIndex < 0 || allItems[i].getIndex() < insertIndex) {
        insertIndex = allItems[i].getIndex();
      }
      form.deleteItem(allItems[i]);
    }
  }

  // ---- 新形式のテキスト欄を作成（既存TEXTなら何もしない） ----
  const newQs = [Q.ORDER_1, Q.ORDER_2, Q.ORDER_3];
  const newItems = [];

  for (const qTitle of newQs) {
    const existing = findItemByTitleOrNull_(form, qTitle);
    if (existing && existing.getType() === FormApp.ItemType.TEXT) continue;

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
