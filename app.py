doc = Document()
    doc.add_heading("Named Range JSON Summary", 0)
    for name, summary in summaries.items():
        doc.add_heading(name, level=1)
        for key, value in summary.items():
            if isinstance(value, (list, dict)):
                value = json.dumps(value, indent=2)
            doc.add_paragraph(f"{key}: {value}")

    docx_io = BytesIO()
    doc.save(docx_io)
    docx_io.seek(0)

    st.download_button(
        "📄 Download Summary as Word Document",
        data=docx_io,
        file_name="named_range_summary.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

