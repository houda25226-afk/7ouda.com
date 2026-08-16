        else:
            df = pd.read_excel(uploaded_file)
    except Exception as e:
        st.error(f"مش قادر أقرأ الملف: {e}")
        st.stop()

    # ------------------------------------------------------
    # معالجة تلقائية للملف بعد رفعه مباشرة:
    #   1) تغيير اسم عمود "Note" لـ "الافادة" (لو موجود)
    #   2) حذف أول صف بيانات بعد صف العناوين (index 0)
    # ------------------------------------------------------
    if "Note" in df.columns:
        df = df.rename(columns={"Note": "الافادة"})

    if len(df) > 0:
        df = df.iloc[1:].reset_index(drop=True)

    st.markdown(
        f'<div class="card">✅ تم تحميل الملف بنجاح — عدد الصفوف: <b>{len(df)}</b></div>',
        unsafe_allow_html=True,
    )
    st.dataframe(df.head(10), use_container_width=True)

    if text_column_input not in df.columns:
        st.error(
            f"عمود '{text_column_input}' مش موجود في الملف. "
            f"الأعمدة الموجودة فعلاً: {', '.join(df.columns.astype(str))}"
        )
        st.stop()

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    if st.button("🚀 ابدأ التصنيف", type="primary", use_container_width=True):
        tokenizer, model, device = load_model()

        texts = df[text_column_input].tolist()
        preds, confidences = predict_batch(texts, tokenizer, model, device)

        result_df = df.copy()
        result_df["التصنيف_المتوقع"] = [LABEL_MAP[p] for p in preds]
        result_df["نسبة_الثقة"] = [round(c * 100, 1) for c in confidences]

        st.success("تم التصنيف بنجاح ✅")
        st.dataframe(result_df, use_container_width=True)

        counts = result_df["التصنيف_المتوقع"].value_counts()
        col1, col2 = st.columns(2)
        col1.metric("✅ إفادات ناجحة", int(counts.get("ناجحة", 0)))
        col2.metric("⛔ إفادات غير ناجحة", int(counts.get("غير ناجحة", 0)))

        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

        if uploaded_file.name.endswith(".csv"):
            output = result_df.to_csv(index=False).encode("utf-8-sig")
            file_name = "نتائج_التصنيف.csv"
            mime = "text/csv"
        else:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                result_df.to_excel(writer, index=False, sheet_name="النتائج")
            output = buffer.getvalue()
            file_name = "نتائج_التصنيف.xlsx"
            mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

        st.download_button(
            label="⬇️ تحميل الملف مع التصنيف",
            data=output,
            file_name=file_name,
            mime=mime,
            use_container_width=True,
        )
else:
    st.markdown(
        '<div class="card" style="text-align:center; color: var(--text-dim);">'
        "📂 ارفع ملف عشان تبدأ — لازم يحتوي على عمود بالنص المراد تصنيفه."
        "</div>",
        unsafe_allow_html=True,
    )
