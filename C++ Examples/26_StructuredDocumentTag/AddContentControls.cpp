#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddContentControls.docx";

	//Creat a new word document.
	Document* document = new Document();
	Section* section = document->AddSection();
	Paragraph* paragraph = section->AddParagraph();
	TextRange* txtRange = paragraph->AppendText(L"The following example shows how to add content controls in a Word document.");
	paragraph = section->AddParagraph();

	//Add Combo Box Content Control.
	paragraph = section->AddParagraph();
	txtRange = paragraph->AppendText(L"Add Combo Box Content Control:  ");
	txtRange->GetCharacterFormat()->SetItalic(true);
	StructureDocumentTagInline* sd = new StructureDocumentTagInline(document);
	paragraph->GetChildObjects()->Add(sd);
	sd->GetSDTProperties()->SetSDTType(SdtType::ComboBox);
	SdtComboBox* cb = new SdtComboBox();
	SdtListItem tempVar(L"Spire.Doc");
	cb->GetListItems()->Add(&tempVar);
	SdtListItem tempVar2(L"Spire.XLS");
	cb->GetListItems()->Add(&tempVar2);
	SdtListItem tempVar3(L"Spire.PDF");
	cb->GetListItems()->Add(&tempVar3);
	sd->GetSDTProperties()->SetControlProperties(cb);
	TextRange* rt = new TextRange(document);
	rt->SetText(cb->GetListItems()->GetItem(0)->GetDisplayText());
	sd->GetSDTContent()->GetChildObjects()->Add(rt);

	section->AddParagraph();

	//Add Text Content Control.
	paragraph = section->AddParagraph();
	txtRange = paragraph->AppendText(L"Add Text Content Control:  ");
	txtRange->GetCharacterFormat()->SetItalic(true);
	sd = new StructureDocumentTagInline(document);
	paragraph->GetChildObjects()->Add(sd);
	sd->GetSDTProperties()->SetSDTType(SdtType::Text);
	SdtText* text = new SdtText(true);
	text->SetIsMultiline(true);
	sd->GetSDTProperties()->SetControlProperties(text);
	rt = new TextRange(document);
	rt->SetText(L"Text");
	sd->GetSDTContent()->GetChildObjects()->Add(rt);

	section->AddParagraph();

	//Add Picture Content Control.
	paragraph = section->AddParagraph();
	txtRange = paragraph->AppendText(L"Add Picture Content Control:  ");
	txtRange->GetCharacterFormat()->SetItalic(true);
	sd = new StructureDocumentTagInline(document);
	paragraph->GetChildObjects()->Add(sd);
	sd->GetSDTProperties()->SetSDTType(SdtType::Picture);
	DocPicture* pic = new DocPicture(document);
	pic->SetWidth(10);
	pic->SetHeight(10);
	pic->LoadImageSpire(Image::FromFile((input_path + L"logo.png").c_str()));
	sd->GetSDTContent()->GetChildObjects()->Add(pic);
	section->AddParagraph();

	//Add Date Picker Content Control.
	paragraph = section->AddParagraph();
	txtRange = paragraph->AppendText(L"Add Date Picker Content Control:  ");
	txtRange->GetCharacterFormat()->SetItalic(true);
	sd = new StructureDocumentTagInline(document);
	paragraph->GetChildObjects()->Add(sd);
	sd->GetSDTProperties()->SetSDTType(SdtType::DatePicker);
	SdtDate* date = new SdtDate();
	date->SetCalendarType(CalendarType::Default);
	date->SetDateFormatSpire(L"yyyy.MM.dd");
	date->SetFullDate(DateTime::GetNow());
	sd->GetSDTProperties()->SetControlProperties(date);
	rt = new TextRange(document);
	rt->SetText(L"1990.02.08");
	sd->GetSDTContent()->GetChildObjects()->Add(rt);

	section->AddParagraph();

	//Add Drop-Down List Content Control.
	paragraph = section->AddParagraph();
	txtRange = paragraph->AppendText(L"Add Drop-Down List Content Control:  ");
	txtRange->GetCharacterFormat()->SetItalic(true);
	sd = new StructureDocumentTagInline(document);
	paragraph->GetChildObjects()->Add(sd);
	sd->GetSDTProperties()->SetSDTType(SdtType::DropDownList);
	SdtDropDownList* sddl = new SdtDropDownList();
	SdtListItem tempVar4(L"Harry");
	sddl->GetListItems()->Add(&tempVar4);
	SdtListItem tempVar5(L"Jerry");
	sddl->GetListItems()->Add(&tempVar5);
	sd->GetSDTProperties()->SetControlProperties(sddl);
	rt = new TextRange(document);
	rt->SetText(sddl->GetListItems()->GetItem(0)->GetDisplayText());
	sd->GetSDTContent()->GetChildObjects()->Add(rt);

	//Save the document.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}