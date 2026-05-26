package com.docxviewer

import android.app.Activity
import android.content.Intent
import android.net.Uri
import android.os.Bundle
import android.widget.Button
import android.widget.TextView
import android.widget.Toast
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import org.apache.poi.xwpf.usermodel.XWPFDocument
import java.io.InputStream

class MainActivity : AppCompatActivity() {

    private lateinit var contentTextView: TextView
    private lateinit var selectButton: Button
    private lateinit var fileNameTextView: TextView

    private val filePickerLauncher = registerForActivityResult(
        ActivityResultContracts.StartActivityForResult()
    ) { result ->
        if (result.resultCode == Activity.RESULT_OK) {
            result.data?.data?.let { uri ->
                readDocxFile(uri)
            }
        }
    }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(R.layout.activity_main)

        contentTextView = findViewById(R.id.contentTextView)
        selectButton = findViewById(R.id.selectButton)
        fileNameTextView = findViewById(R.id.fileNameTextView)

        selectButton.setOnClickListener {
            openFilePicker()
        }

        if (intent?.action == Intent.ACTION_VIEW) {
            intent.data?.let { uri ->
                readDocxFile(uri)
            }
        }
    }

    private fun openFilePicker() {
        val intent = Intent(Intent.ACTION_GET_CONTENT).apply {
            type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            addCategory(Intent.CATEGORY_OPENABLE)
        }
        filePickerLauncher.launch(Intent.createChooser(intent, "Select DOCX File"))
    }

    private fun readDocxFile(uri: Uri) {
        try {
            val inputStream: InputStream? = contentResolver.openInputStream(uri)
            inputStream?.use { stream ->
                val document = XWPFDocument(stream)
                val content = StringBuilder()
                
                fileNameTextView.text = uri.lastPathSegment ?: "Document"
                
                content.append("=== Document Content ===\n\n")
                
                for (paragraph in document.paragraphs) {
                    val text = paragraph.text
                    if (text.isNotBlank()) {
                        content.append(text).append("\n\n")
                    }
                }
                
                contentTextView.text = content.toString()
            }
        } catch (e: Exception) {
            Toast.makeText(this, "Error: ${e.message}", Toast.LENGTH_LONG).show()
            contentTextView.text = "Error reading document: ${e.message}"
        }
    }
}
