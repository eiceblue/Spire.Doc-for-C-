#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"MultiStylesInAParagraph.docx";
	
	//Create a Word document
	Document* doc = new Document();

	//Add a section
	Section* section = doc->AddSection();

	//Add a paragraph
	Paragraph* para = section->AddParagraph();

	//Add a text range 1 and set its style
	TextRange* range = para->AppendText(L"Spire.Doc for C++");
	range->GetCharacterFormat()->SetFontName(L"Calibri");
	range->GetCharacterFormat()->SetFontSize(16.0f);
	range->GetCharacterFormat()->SetTextColor(Spire::Common::Color::GetBlue());
	range->GetCharacterFormat()->SetBold(true);
	range->GetCharacterFormat()->SetUnderlineStyle(UnderlineStyle::Single);

	//Add a text range 2 and set its style
	range = para->AppendText(L"is a professional Word C++ library");
	range->GetCharacterFormat()->SetFontName(L"Calibri");
	range->GetCharacterFormat()->SetFontSize(15.0f);

	//Save the Word document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
	delete doc;
}