#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main(){
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"BlankTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"SetHyperlinkFormat.docx";

	//Load Document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());
	Section* section = doc->GetSections()->GetItem(0);

	//Add a paragraph and append a hyperlink to the paragraph
	Paragraph* para1 = section->AddParagraph();
	para1->AppendText(L"Regular Link: ");
	//Format the hyperlink with default color and underline style
	TextRange* txtRange1 = para1->AppendHyperlink(L"www.e-iceblue.com", L"www.e-iceblue.com", HyperlinkType::WebLink);
	txtRange1->GetCharacterFormat()->SetFontName(L"Times New Roman");
	txtRange1->GetCharacterFormat()->SetFontSize(12);
	Paragraph* blankPara1 = section->AddParagraph();

	//Add a paragraph and append a hyperlink to the paragraph
	Paragraph* para2 = section->AddParagraph();
	para2->AppendText(L"Change Color: ");
	//Format the hyperlink with red color and underline style
	TextRange* txtRange2 = para2->AppendHyperlink(L"www.e-iceblue.com", L"www.e-iceblue.com", HyperlinkType::WebLink);
	txtRange2->GetCharacterFormat()->SetFontName(L"Times New Roman");
	txtRange2->GetCharacterFormat()->SetFontSize(12);
	txtRange2->GetCharacterFormat()->SetTextColor(Color::GetRed());
	Paragraph* blankPara2 = section->AddParagraph();

	//Add a paragraph and append a hyperlink to the paragraph
	Paragraph* para3 = section->AddParagraph();
	para3->AppendText(L"Remove Underline: ");
	//Format the hyperlink with red color and no underline style
	TextRange* txtRange3 = para3->AppendHyperlink(L"www.e-iceblue.com", L"www.e-iceblue.com", HyperlinkType::WebLink);
	txtRange3->GetCharacterFormat()->SetFontName(L"Times New Roman");
	txtRange3->GetCharacterFormat()->SetFontSize(12);
	txtRange3->GetCharacterFormat()->SetUnderlineStyle(UnderlineStyle::None);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}