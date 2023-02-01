#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"PageBorderSurround.docx";

	//Create a new document
	Document* doc = new Document();
	Section* section = doc->AddSection();

	//Add a sample page border to the document
	section->GetPageSetup()->GetBorders()->SetBorderType(BorderStyle::Wave);
	section->GetPageSetup()->GetBorders()->SetColor(Color::GetGreen());
	section->GetPageSetup()->GetBorders()->GetLeft()->SetSpace(20);
	section->GetPageSetup()->GetBorders()->GetRight()->SetSpace(20);

	//Add a header and set its format
	Paragraph* paragraph1 = section->GetHeadersFooters()->GetHeader()->AddParagraph();
	paragraph1->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Right);
	TextRange* headerText = paragraph1->AppendText(L"Header isn't included in page border");
	headerText->GetCharacterFormat()->SetFontName(L"Calibri");
	headerText->GetCharacterFormat()->SetFontSize(20);
	headerText->GetCharacterFormat()->SetBold(true);

	//Add a footer and set its format
	Paragraph* paragraph2 = section->GetHeadersFooters()->GetFooter()->AddParagraph();
	paragraph2->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Left);
	TextRange* footerText = paragraph2->AppendText(L"Footer is included in page border");
	footerText->GetCharacterFormat()->SetFontName(L"Calibri");
	footerText->GetCharacterFormat()->SetFontSize(20);
	footerText->GetCharacterFormat()->SetBold(true);

	//Set the header not included in the page border while the footer included
	section->GetPageSetup()->SetPageBorderIncludeHeader(false);
	section->GetPageSetup()->SetHeaderDistance(40);
	section->GetPageSetup()->SetPageBorderIncludeFooter(true);
	section->GetPageSetup()->SetFooterDistance(40);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
