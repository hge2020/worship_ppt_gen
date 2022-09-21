from Lyrics_gen import GenLyrics
from Ad_gen import GenAd
from Bible_gen import GenBible, GenBibleAdditional
from Any_gen import GenAny

start_file = 'template.pptx'
save_file = '220925_청년부.pptx'

GenLyrics(start_file, save_file, '가사.txt')
GenAd(save_file, save_file)
GenBible(save_file, save_file)
# GenAny(save_file, save_file, '성경_추가.txt')
# GenBibleAdditional(save_file, save_file, '성경_추가.txt')
GenLyrics(save_file, save_file,  '결찬가사.txt')