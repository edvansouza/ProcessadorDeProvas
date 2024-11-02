# ibe-csv-to-excel
Converte arquivos CSV do Google Forms para uma Tabela de Excel


Este um projeto para converter Google Forms com sessões em tabela do excel.
Provas são feitas numa escola regular pelo Google Forms e as matérias são separadas por sessões.
Cada professor de cada matéria precisa coletar as notas dos alunos mas o Google Forms não gera a pontuação separada de cada sessão apenas a pontuação total somando todas as questões de todas as matérias.
Desta forma cada professor de cada matéria precisa verificar a pontuação de sua matéria aluno por aluno, turma por turma.
O código python separa todas mas matérias em subtabelas numa arquivo de Excel e dá a pontuação total aluno por aluno, turma por turma a partir do envio do arquivo csv gerado do resultado do Google Forms.

O código foi hospedo no streamlit no endereço:

`https://gforms-csv-to-excel.streamlit.app`
