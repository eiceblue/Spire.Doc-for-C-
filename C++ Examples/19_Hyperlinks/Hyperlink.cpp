#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"Hyperlink.docx";

	//Open a blank word document as template
	Document* document = new Document();
	Section* section = document->AddSection();

	//Insert hyperlink
	InsertHyperlink(section);

	//Save doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

void InsertHyperlink(Section* section)
{
	Paragraph* paragraph = section->GetParagraphs()->GetCount() > 0 ? section->GetParagraphs()->GetItem(0) : section->AddParagraph();
	paragraph->AppendText(L"Spire.Doc for C++ \r\n e-iceblue company Ltd. 2002-2010 All rights reserverd");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);

	paragraph = section->AddParagraph();
	paragraph->AppendText(L"Home page");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);
	paragraph = section->AddParagraph();
	paragraph->AppendHyperlink(L"www.e-iceblue.com", L"www.e-iceblue.com", HyperlinkType::WebLink);

	paragraph = section->AddParagraph();
	paragraph->AppendText(L"Contact US");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);
	paragraph = section->AddParagraph();
	paragraph->AppendHyperlink(L"mailto:support@e-iceblue.com", L"support@e-iceblue.com", HyperlinkType::EMailLink);

	paragraph = section->AddParagraph();
	paragraph->AppendText(L"Forum");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);
	paragraph = section->AddParagraph();
	paragraph->AppendHyperlink(L"www.e-iceblue.com/forum/", L"www.e-iceblue.com/forum/", HyperlinkType::WebLink);

	paragraph = section->AddParagraph();
	paragraph->AppendText(L"Download Link");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);
	paragraph = section->AddParagraph();
	paragraph->AppendHyperlink(L"www.e-iceblue.com/Download/download-word-for-net-now.html", L"www.e-iceblue.com/Download/download-word-for-net-now.html", HyperlinkType::WebLink);

	paragraph = section->AddParagraph();
	paragraph->AppendText(L"Insert Link On Image");
	paragraph->ApplyStyle(BuiltinStyle::Heading2);
	paragraph = section->AddParagraph();
	wstring input_path = DATAPATH;
	wstring imagePath = input_path + L"Spire.Doc.png";
	DocPicture* picture = paragraph->AppendPicture(Image::FromFile(imagePath.c_str()));
	paragraph->AppendHyperlink(L"www.e-iceblue.com/Download/download-word-for-net-now.html", picture, HyperlinkType::WebLink);
}
