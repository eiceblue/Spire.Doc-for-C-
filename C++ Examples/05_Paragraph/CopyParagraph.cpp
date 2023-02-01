#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_Docx_5.docx";
	wstring logoFile = input_path + L"Logo.jpg";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CopyParagraph.docx";

	//Create Word document1.
	Document* document1 = new Document();

	//Load the file from disk.
	document1->LoadFromFile(inputFile.c_str());

	//Create a new document.
	Document* document2 = new Document();

	//Get paragraph 1 and paragraph 2 in document1.
	Section* s = document1->GetSections()->GetItem(0);
	Paragraph* p1 = s->GetParagraphs()->GetItem(0);
	Paragraph* p2 = s->GetParagraphs()->GetItem(1);

	//Copy p1 and p2 to document2.
	Section* s2 = document2->AddSection();
	Paragraph* NewPara1 = dynamic_cast<Paragraph*>(p1->Clone());
	s2->GetParagraphs()->Add(NewPara1);
	Paragraph* NewPara2 = dynamic_cast<Paragraph*>(p2->Clone());
	s2->GetParagraphs()->Add(NewPara2);

	//Add watermark.
	PictureWatermark* WM = new PictureWatermark();
	WM->SetPicture(logoFile.c_str());

	document2->SetWatermark(WM);

	//Save the file.
	document2->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	document1->Close();
	document2->Close();
	delete document1;
	delete document2;
}