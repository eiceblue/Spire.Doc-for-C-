#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"TextWaterMark.docx";

	//Open a Word document as template.
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Insert text watermark.
	InsertTextWatermark(document->GetSections()->GetItem(0));
	//Save as docx file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

void InsertTextWatermark(Section* section)
{
	TextWatermark* txtWatermark = new TextWatermark();
	txtWatermark->SetText(L"E-iceblue");
	txtWatermark->SetFontSize(95);
	txtWatermark->SetColor(Color::GetBlue());
	txtWatermark->SetLayout(WatermarkLayout::Diagonal);
	section->GetDocument()->SetWatermark(txtWatermark);
}
