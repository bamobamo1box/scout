import pandas as pd
import streamlit as st
import os

def load_data():
    """データを読み込む関数"""
    possible_files = ["data.xlsx", "data.xlsx.xlsx"]
    file_path = None

    for file in possible_files:
        if os.path.exists(file):
            file_path = file
            break

    if not file_path:
        st.error(f"ファイルが見つかりません: {', '.join(possible_files)} のいずれかが必要です。")
        return None

    try:
        data = pd.read_excel(file_path, engine="openpyxl")
        data.columns = data.columns.str.strip()
        data = data.fillna("")

        # データが363行目までしかデータがない場合、それ以降を削除
        data = data.iloc[:363]

        for col in data.select_dtypes(include=["datetime"]).columns:
            data[col] = data[col].apply(lambda x: x.strftime("%Y/%m/%d") if not pd.isnull(x) else "")

        required_columns = ["マッチング", "書類通過", "承諾"]
        missing_columns = [col for col in required_columns if col not in data.columns]
        if missing_columns:
            st.error(f"必要な列がデータに存在しません: {', '.join(missing_columns)}")
            return None

        for col in required_columns:
            data[col] = data[col].replace({"True": 1, "False": 0, "": 0}).astype(int)

        return data
    except Exception as e:
        st.error(f"データの読み込み中にエラーが発生しました: {e}")
        return None

def initialize_session_state(data):
    """セッション状態の初期化"""
    if "quiz_data" not in st.session_state:
        if data is not None and not data.empty:
            st.session_state.quiz_data = data.sample(n=min(30, len(data)), random_state=None).reset_index(drop=True)
            st.session_state.current_index = 0
            st.session_state.current_step = "マッチング"
            st.session_state.answers = []
            st.session_state.show_next_button = False
            st.session_state.answer_submitted = False  # 新規フラグ追加
        else:
            st.error("クイズデータの初期化に失敗しました。データが空か不正です。")

def export_results_to_excel(results):
    """結果をExcelファイルに出力する関数"""
    output_file = "quiz_results.xlsx"
    results_df = pd.DataFrame(results)
    results_df.to_excel(output_file, index=False)
    return output_file

def apply_custom_styles():
    """カスタムスタイルを適用する関数"""
    st.markdown(
        """
        <style>
        body {
            background-color: #FAE400;
            color: #333;
            font-size: 16px; /* ベースフォントサイズ */
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        .stButton>button {
            background-color: #FAE400;
            color: #000;
            border: none;
            border-radius: 5px;
            padding: 12px 20px;
            font-size: 16px;
            cursor: pointer;
            width: 100%; /* モバイル向けに幅を調整 */
        }
        .stButton>button:hover {
            background-color: #FFD700;
        }
        .stRadio>div>div {
            background-color: #FFF9C4;
            border-radius: 5px;
            padding: 10px;
            margin-bottom: 10px;
            font-size: 16px;
        }
        .stRadio>div>div:hover {
            background-color: #FFE082;
        }
        .stMarkdown {
            background-color: #FAF3B0;
            border-radius: 5px;
            padding: 10px;
            font-size: 16px;
            color: #333;
        }
        @media only screen and (max-width: 600px) {
            body {
                font-size: 14px; /* モバイル用フォントサイズ */
            }
            .stButton>button {
                font-size: 14px;
                padding: 10px 16px;
            }
            .stRadio>div>div {
                font-size: 14px;
                padding: 8px;
            }
        }
        </style>
        """,
        unsafe_allow_html=True
    )

def format_question_text(step):
    """ステップごとの質問文を整形する関数"""
    step_mapping = {
        "マッチング": "この求職者とマッチングしましたか？",
        "書類通過": "この求職者の書類は通過しましたか？",
        "承諾": "この求職者は内定を承諾しましたか？"
    }
    return step_mapping.get(step, "このステップを完了しましたか？")

def format_displayed_data(data):
    """データをMarkdownでフォーマットする関数"""
    formatted_data = """
    <div class="stMarkdown">
    <ul>
    """
    for key, value in data.items():
        formatted_data += f"<li><b>{key}:</b> {value}</li>\n"
    formatted_data += """</ul>
    </div>
    """
    return formatted_data

def main():
    st.title("ミキワメクイズ")

    # ロゴを表示
    logo_path = "logo.png"  # ロゴ画像のパス
    if os.path.exists(logo_path):
        st.image(logo_path, width=200)
    else:
        st.warning("ロゴ画像が見つかりません")

    apply_custom_styles()

    data = load_data()
    if data is None or data.empty:
        st.error("データが正しく読み込まれませんでした。正しいExcelファイルを配置してください。")
        return

    initialize_session_state(data)

    if "quiz_data" not in st.session_state or st.session_state.quiz_data.empty:
        st.error("クイズデータが正しく初期化されませんでした。")
        return

    quiz_data = st.session_state.quiz_data
    current_index = st.session_state.current_index

    if current_index < len(quiz_data):
        row = quiz_data.iloc[current_index]

        st.subheader(f"求職者 {current_index + 1} / {len(quiz_data)}")
        displayed_data = row.drop(["マッチング", "書類通過", "承諾", "入社企業"], errors="ignore").to_dict()

        if not any(displayed_data.values()):
            st.warning("求職者データが不完全です。次へ進みます。")
            st.session_state.current_index += 1
            st.experimental_rerun()
            return

        st.markdown(format_displayed_data(displayed_data), unsafe_allow_html=True)

        question_text = format_question_text(st.session_state.current_step)
        answer = st.radio(
            question_text,
            ["YES", "NO"],
            key=f"{current_index}_{st.session_state.current_step}"
        )

        if st.button("回答する", key=f"answer_{current_index}_{st.session_state.current_step}"):
            correct_answer = "YES" if row[st.session_state.current_step] == 1 else "NO"

            if correct_answer == "YES":
                if answer == "YES":
                    st.success("正解です！次のステップへ進みます。")
                else:
                    st.error("不正解です。次のステップに進みます。")

                next_step = {"マッチング": "書類通過", "書類通過": "承諾"}
                if st.session_state.current_step == "承諾":
                    st.session_state.current_index += 1
                    st.session_state.current_step = "マッチング"
                else:
                    st.session_state.current_step = next_step[st.session_state.current_step]

            else:  # correct_answer == "NO"
                if answer == "YES":
                    st.error("不正解です。次のステップに進みます。")
                else:
                    st.success("正解です！次のステップに進みます。")
                st.session_state.current_index += 1
                st.session_state.current_step = "マッチング"

            st.session_state.answers.append({
                "求職者": current_index + 1,
                "ステップ": st.session_state.current_step,
                "回答": answer,
                "キャリアサマリ": row.get("キャリアサマリ", ""),
                "正解/不正解": "正解" if answer == correct_answer else "不正解"
            })
            st.session_state.show_next_button = True
            st.session_state.answer_submitted = True  # 回答が行われたことを記録

        if st.session_state.show_next_button:
            if st.button("次へ", key=f"next_{current_index}"):
                if st.session_state.answer_submitted:
                    st.session_state.show_next_button = False
                    st.session_state.answer_submitted = False  # フラグをリセット
                    st.experimental_rerun()
                else:
                    st.warning("回答するボタンを押してください。")

    else:
        st.write("すべてのクイズが終了しました！")

        # 結果の集計
        results = st.session_state.answers
        results_df = pd.DataFrame(results)

        correct_count = results_df[results_df["正解/不正解"] == "正解"].shape[0]
        incorrect_count = results_df[results_df["正解/不正解"] == "不正解"].shape[0]
        total_questions = len(results)
        accuracy_rate = (correct_count / total_questions) * 100 if total_questions > 0 else 0

        st.write(f"正解数: {correct_count}")
        st.write(f"不正解数: {incorrect_count}")
        st.write(f"正答率: {accuracy_rate:.2f}%")

        # 結果の表示
        st.write("詳細結果:")
        st.dataframe(results_df)

        # Excelファイルのダウンロード
        output_file = export_results_to_excel(results)
        with open(output_file, "rb") as f:
            st.download_button(
                label="結果をダウンロード (Excel)",
                data=f,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
