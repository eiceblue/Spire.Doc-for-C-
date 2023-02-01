#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"TextBoxFormat.docx";

	//Create a new document
	Document* doc = new Document();
	Section* sec = doc->AddSection();

	//Add a text box and append sample text
	TextBox* TB = doc->GetSections()->GetItem(0)->AddParagraph()->AppendTextBox(310, 90);
	Paragraph* para = TB->GetBody()->AddParagraph();
	TextRange* TR = para->AppendText(L"Using Spire.Doc, developers will find a simple and effective method to endow their applications with rich MS Word features. ");
	TR->GetCharacterFormat()->SetFontName(L"Cambria ");
	TR->GetCharacterFormat()->SetFontSize(13);

	//Set exact position for the text box
	TB->GetFormat()->SetHorizontalOrigin(HorizontalOrigin::Page);
	TB->GetFormat()->SetHorizontalPosition(120);
	TB->GetFormat()->SetVerticalOrigin(VerticalOrigin::Page);
	TB->GetFormat()->SetVerticalPosition(100);

	//Set line style for the text box
	TB->GetFormat()->SetLineStyle(TextBoxLineStyle::Double);
	TB->GetFormat()->SetLineColor(Spire::Common::Color::GetCornflowerBlue());
	TB->GetFormat()->SetLineDashing(LineDashing::Solid);
	TB->GetFormat()->SetLineWidth(5);

	//Set internal margin for the text box
	TB->GetFormat()->GetInternalMargin()->SetTop(15);
	TB->GetFormat()->GetInternalMargin()->SetBottom(10);
	TB->GetFormat()->GetInternalMargin()->SetLeft(12);
	TB->GetFormat()->GetInternalMargin()->SetRight(10);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
