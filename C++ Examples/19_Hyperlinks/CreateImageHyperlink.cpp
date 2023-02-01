#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"BlankTemplate.docx";
	wstring inputFile_1 = input_path + L"Spire.Doc.png";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CreateImageHyperlink.docx";

	//Load Document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	Section* section = doc->GetSections()->GetItem(0);
	//Add a paragraph
	Paragraph* paragraph = section->AddParagraph();
	//Load an image to a DocPicture object

	Image* image = Image::FromFile(inputFile_1.c_str());
	DocPicture* picture = new DocPicture(doc);
	//Add an image hyperlink to the paragraph
	picture->LoadImageSpire(image);
	paragraph->AppendHyperlink(L"https://www.e-iceblue.com/Introduce/word-for-net-introduce.html", picture, HyperlinkType::WebLink);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
