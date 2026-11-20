from pptx import Presentation
prs = Presentation(r'D:\Preksha_carelon\Carelon_CaseStudy_Presentation.pptx')
print(f'Total slides: {len(prs.slides)}')
print(f'Slide dimensions: {prs.slide_width.inches:.2f}" x {prs.slide_height.inches:.2f}"')
for i, slide in enumerate(prs.slides):
    notes = slide.notes_slide.notes_text_frame.text
    has_notes = bool(notes.strip())
    print(f'  Slide {i+1}: has_notes={has_notes} | notes_len={len(notes)}')
print('VERIFICATION COMPLETE')
