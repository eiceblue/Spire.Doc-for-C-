#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template.docx";
	wstring imagePath1 = input_path + L"E-iceblue.png";
	wstring imagePath2 = input_path + L"logo.png";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ImageHeaderAndFooter.docx";

	//Load the document from disk
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the header of the first page
	HeaderFooter* header = doc->GetSections()->GetItem(0)->GetHeadersFooters()->GetHeader();

	//Add a paragraph for the header
	Paragraph* paragraph = header->AddParagraph();

	//Set the format of the paragraph
	paragraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);

	//Append a picture in the paragraph
	DocPicture* headerimage = paragraph->AppendPicture(Image::FromFile(imagePath1.c_str()));
	headerimage->SetVerticalAlignment(ShapeVerticalAlignment::Bottom);

	//Get the footer of the first section
	HeaderFooter* footer = doc->GetSections()->GetItem(0)->GetHeadersFooters()->GetFooter();

	//Add a paragraph for the footer
	Paragraph* paragraph2 = footer->AddParagraph();

	//Set the format of the paragraph
	paragraph2->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Left);

	//Append a picture in the paragraph
	DocPicture* footerimage = paragraph2->AppendPicture(Image::FromFile(imagePath2.c_str()));
	//Append text in the paragraph
	wstring string = L"Copyright \u00A9 2013 e-iceblue. All Rights Reserved.";
	TextRange* TR = paragraph2->AppendText(string.c_str());
	TR->GetCharacterFormat()->SetFontName(L"Arial");
	TR->GetCharacterFormat()->SetFontSize(10);
	TR->GetCharacterFormat()->SetTextColor(Color::GetBlack());

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
