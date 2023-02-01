#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"MultiplePages.docx";
	wstring imagePath = input_path + L"E-iceblue.png";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"DifferentFirstPage.docx";

	//Load the document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the section and set the property true
	Section* section = doc->GetSections()->GetItem(0);
	section->GetPageSetup()->SetDifferentFirstPageHeaderFooter(true);

	//Set the first page header. Here we append a picture in the header
	Paragraph* paragraph1 = section->GetHeadersFooters()->GetFirstPageHeader()->AddParagraph();
	paragraph1->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);
	DocPicture* headerimage = paragraph1->AppendPicture(Image::FromFile(imagePath.c_str()));
	//Set the first page footer
	Paragraph* paragraph2 = section->GetHeadersFooters()->GetFirstPageFooter()->AddParagraph();
	paragraph2->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
	TextRange* FF = paragraph2->AppendText(L"First Page Footer");
	FF->GetCharacterFormat()->SetFontSize(10);

	//Set the other header & footer. If you only need the first page header & footer, don't set this
	Paragraph* paragraph3 = section->GetHeadersFooters()->GetHeader()->AddParagraph();
	paragraph3->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
	TextRange* NH = paragraph3->AppendText(L"Spire.Doc for C++");
	NH->GetCharacterFormat()->SetFontSize(10);

	Paragraph* paragraph4 = section->GetHeadersFooters()->GetFooter()->AddParagraph();
	paragraph4->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Center);
	TextRange* NF = paragraph4->AppendText(L"E-iceblue");
	NF->GetCharacterFormat()->SetFontSize(10);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
