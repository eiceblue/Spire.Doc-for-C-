#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"BlankTemplate.docx";
	wstring imagePath = input_path + L"Word.png";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertImage.docx";

	//Load Document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	Section* section = doc->GetSections()->GetItem(0);
	Paragraph* paragraph = section->GetParagraphs()->GetCount() > 0 ? section->GetParagraphs()->GetItem(0) : section->AddParagraph();
	paragraph->AppendText(L"The sample demonstrates how to insert an image into a document.");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);
	paragraph = section->AddParagraph();
	paragraph->AppendText(L"The above is a picture.");

	//Create a picture
	DocPicture* picture = new DocPicture(doc);
	picture->LoadImageSpire(imagePath.c_str());

	//set image's position
	picture->SetHorizontalPosition(50.0F);
	picture->SetVerticalPosition(60.0F);

	//set image's size
	picture->SetWidth(200);
	picture->SetHeight(200);

	//set textWrappingStyle with image;
	picture->SetTextWrappingStyle(TextWrappingStyle::Through);
	//Insert the picture at the beginning of the second paragraph
	paragraph->GetChildObjects()->Insert(0, picture);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
