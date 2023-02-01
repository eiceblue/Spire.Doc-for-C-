#include "pch.h"
#include "TextRangeLocation.h"

using namespace Spire::Doc;

namespace DocTest {
	namespace Demo {
		namespace _02_FindAndReplace
		{
			TextRangeLocation::TextRangeLocation(TextRange* text)
			{
				this->SetText(text);
			}

			TextRange* TextRangeLocation::GetText()
			{
				return m_Text;
			}

			void TextRangeLocation::SetText(TextRange* value)
			{
				m_Text = value;
			}

			Paragraph* TextRangeLocation::GetOwner()
			{
				return this->GetText()->GetOwnerParagraph();
			}

			int TextRangeLocation::GetIndex()
			{
				return this->GetOwner()->GetChildObjects()->IndexOf(this->GetText());
			}

			int TextRangeLocation::CompareTo(TextRangeLocation* other)
			{
				return -(this->GetIndex() - other->GetIndex());
			}
		}
	}
}