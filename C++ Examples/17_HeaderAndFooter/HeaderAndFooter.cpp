#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Sample.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"HeaderAndFooter.docx";

	//Create word document
	Document* document = new Document();

	document->LoadFromFile(inputFile.c_str());
	Section* section = document->GetSections()->GetItem(0);

	//insert header and footer
	InsertHeaderAndFooter(section);

	//Save as docx file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

void InsertHeaderAndFooter(Section* section)
{
	HeaderFooter* header = section->GetHeadersFooters()->GetHeader();
	HeaderFooter* footer = section->GetHeadersFooters()->GetFooter();

	//insert picture and text to header
	Paragraph* headerParagraph = header->AddParagraph();
	wstring input_path = DATAPATH;
	wstring imagePath1 = input_path + L"Header.png";
	DocPicture* headerPicture = headerParagraph->AppendPicture(Image::FromFile(imagePath1.c_str()));
	//header text
	TextRange* text = headerParagraph->AppendText(L"Demo of Spire.Doc");
	text->GetCharacterFormat()->SetFontName(L"Arial");
	text->GetCharacterFormat()->SetFontSize(10);
	text->GetCharacterFormat()->SetItalic(true);
	headerParagraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);

	//border
	headerParagraph->GetFormat()->GetBorders()->GetBottom()->SetBorderType(BorderStyle::Single);
	headerParagraph->GetFormat()->GetBorders()->GetBottom()->SetSpace(0.05F);


	//header picture layout - text wrapping
	headerPicture->SetTextWrappingStyle(TextWrappingStyle::Behind);

	//header picture layout - position
	headerPicture->SetHorizontalOrigin(HorizontalOrigin::Page);
	headerPicture->SetHorizontalAlignment(ShapeHorizontalAlignment::Left);
	headerPicture->SetVerticalOrigin(VerticalOrigin::Page);
	headerPicture->SetVerticalAlignment(ShapeVerticalAlignment::Top);

	//insert picture to footer
	Paragraph* footerParagraph = footer->AddParagraph();
	wstring imagePath2 = input_path + L"Footer.png";
	DocPicture* footerPicture = footerParagraph->AppendPicture(Image::FromFile(imagePath2.c_str()));
	//footer picture layout
	footerPicture->SetTextWrappingStyle(TextWrappingStyle::Behind);
	footerPicture->SetHorizontalOrigin(HorizontalOrigin::Page);
	footerPicture->SetHorizontalAlignment(ShapeHorizontalAlignment::Left);
	footerPicture->SetVerticalOrigin(VerticalOrigin::Page);
	footerPicture->SetVerticalAlignment(ShapeVerticalAlignment::Bottom);

	//insert page number
	footerParagraph->AppendField(L"page number", FieldType::FieldPage);
	footerParagraph->AppendText(L" of ");
	footerParagraph->AppendField(L"number of pages", FieldType::FieldNumPages);
	footerParagraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);

	//border
	footerParagraph->GetFormat()->GetBorders()->GetTop()->SetBorderType(BorderStyle::Single);
	footerParagraph->GetFormat()->GetBorders()->GetTop()->SetSpace(0.05F);
}
