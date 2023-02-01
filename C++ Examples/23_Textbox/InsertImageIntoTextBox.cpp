#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertImageIntoTextBox.docx";

	//Create a new document
	Document* doc = new Document();

	Section* section = doc->AddSection();
	Paragraph* paragraph = section->AddParagraph();

	//Append a textbox to paragraph
	TextBox* tb = paragraph->AppendTextBox(220, 220);

	//Set the position of the textbox
	tb->GetFormat()->SetHorizontalOrigin(HorizontalOrigin::Page);
	tb->GetFormat()->SetHorizontalPosition(50);
	tb->GetFormat()->SetVerticalOrigin(VerticalOrigin::Page);
	tb->GetFormat()->SetVerticalPosition(50);

	//Set the fill effect of textbox as picture
	tb->GetFormat()->GetFillEfects()->SetType(BackgroundType::Picture);

	//Fill the textbox with a picture
	wstring input_path = DATAPATH;
	wstring imagePath = input_path + L"Spire.Doc.png";
	tb->GetFormat()->GetFillEfects()->SetPicture(Image::FromFile(imagePath.c_str()));
	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
