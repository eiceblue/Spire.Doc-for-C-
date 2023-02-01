#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SampleB_2.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddImageToEachPage.docx";

	//Open a Word document
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	wstring imgPath = input_path + L"Spire.Doc.png";

	//Add a picture in footer and set it's position
	DocPicture* picture = document->GetSections()->GetItem(0)->GetHeadersFooters()->GetFooter()->AddParagraph()->AppendPicture(Image::FromFile(imgPath.c_str()));
	picture->SetVerticalOrigin(VerticalOrigin::Page);
	picture->SetHorizontalOrigin(HorizontalOrigin::Page);
	picture->SetVerticalAlignment(ShapeVerticalAlignment::Bottom);
	picture->SetTextWrappingStyle(TextWrappingStyle::None);

	//Add a textbox in footer and set it's positiion
	TextBox* textbox = document->GetSections()->GetItem(0)->GetHeadersFooters()->GetFooter()->AddParagraph()->AppendTextBox(150, 20);
	textbox->SetVerticalOrigin(VerticalOrigin::Page);
	textbox->SetHorizontalOrigin(HorizontalOrigin::Page);
	textbox->SetHorizontalPosition(300);
	textbox->SetVerticalPosition(700);
	textbox->GetBody()->AddParagraph()->AppendText(L"Welcome to E-iceblue");

	//Save to file
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}