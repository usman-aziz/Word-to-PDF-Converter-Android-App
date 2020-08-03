package com.example.wordtopdfconverter;

import androidx.annotation.RequiresApi;
import androidx.appcompat.app.AppCompatActivity;

import android.annotation.TargetApi;
import android.app.Activity;
import android.content.Intent;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import com.aspose.words.Document;
import com.aspose.words.License;
import com.github.barteksc.pdfviewer.PDFView;
import com.google.android.material.floatingactionbutton.FloatingActionButton;

import android.os.Environment;
import android.view.View;
import android.widget.TextView;
import android.widget.Toast;

@TargetApi(Build.VERSION_CODES.FROYO)
public class MainActivity extends AppCompatActivity {

    private static final int PICK_PDF_FILE = 2;
    private final String storageDir = Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS) + File.separator;
    private final String outputPDF = storageDir + "Converted_PDF.pdf";
    private TextView textView = null;
    private Uri document = null;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        // apply the license if you have the Aspose.Words license...
        applyLicense();
        // get treeview and set its text
        textView = (TextView) findViewById(R.id.textView);
        textView.setText("Select a Word DOCX file...");
        // define click listener of floating button
        FloatingActionButton myFab = (FloatingActionButton) findViewById(R.id.fab);
        myFab.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                try {
                    // open Word file from file picker and convert to PDF
                    openaAndConvertFile(null);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        });
    }

    private void openaAndConvertFile(Uri pickerInitialUri) {
        // create a new intent to open document
        Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
        intent.addCategory(Intent.CATEGORY_OPENABLE);
        // mime types for MS Word documents
        String[] mimetypes = {"application/vnd.openxmlformats-officedocument.wordprocessingml.document", "application/msword"};
        intent.setType("*/*");
        intent.putExtra(Intent.EXTRA_MIME_TYPES, mimetypes);
        // start activiy
        startActivityForResult(intent, PICK_PDF_FILE);
    }

    @RequiresApi(api = Build.VERSION_CODES.KITKAT)
    @Override
    public void onActivityResult(int requestCode, int resultCode,
                                 Intent intent) {
        super.onActivityResult(requestCode, resultCode, intent);

        if (resultCode == Activity.RESULT_OK) {
            if (intent != null) {
                document = intent.getData();
                // open the selected document into an Input stream
                try (InputStream inputStream =
                             getContentResolver().openInputStream(document);) {
                    Document doc = new Document(inputStream);
                    // save DOCX as PDF
                    doc.save(outputPDF);
                    // show PDF file location in toast as well as treeview (optional)
                    Toast.makeText(MainActivity.this, "File saved in: " + outputPDF, Toast.LENGTH_LONG).show();
                    textView.setText("PDF saved at: " + outputPDF);
                    // view converted PDF
                    viewPDFFile();
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                    Toast.makeText(MainActivity.this, "File not found: " + e.getMessage(), Toast.LENGTH_LONG).show();
                } catch (IOException e) {
                    e.printStackTrace();
                    Toast.makeText(MainActivity.this, e.getMessage(), Toast.LENGTH_LONG).show();
                } catch (Exception e) {
                    e.printStackTrace();
                    Toast.makeText(MainActivity.this, e.getMessage(), Toast.LENGTH_LONG).show();
                }
            }
        }
    }

    public void viewPDFFile() {
        // load PDF into the PDFView
        PDFView pdfView = (PDFView) findViewById(R.id.pdfView);
        pdfView.fromFile(new File(outputPDF)).load();
    }
    public void applyLicense()
    {
        // set license
        //License lic= new License();
        // add a raw resource directory within res and then add your license file as resource...
        //InputStream inputStream = getResources().openRawResource(R.raw.license);
        try {
            //lic.setLicense(inputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}