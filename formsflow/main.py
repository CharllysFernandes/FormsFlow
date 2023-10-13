import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QProgressBar, QPushButton, QFileDialog, QVBoxLayout, QWidget

class FormsFlowApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("FormsFlow")

        # Crie um widget central para a janela
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)

        # Crie um layout vertical para organizar os widgets
        layout = QVBoxLayout(central_widget)

        self.progress = QProgressBar()
        layout.addWidget(self.progress)

        self.btn_enviar_chamados = QPushButton("Enviar Chamados")
        self.btn_enviar_chamados.clicked.connect(self.enviar_chamados)
        layout.addWidget(self.btn_enviar_chamados)

        self.btn_abrir_planilha = QPushButton("Abrir Planilha Chamado")
        self.btn_abrir_planilha.clicked.connect(self.abrir_planilha_chamado)
        layout.addWidget(self.btn_abrir_planilha)

    def enviar_chamados(self):
        # Adicione aqui a lógica para enviar chamados
        pass

    def abrir_planilha_chamado(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Abrir Planilha Chamado", "", "Planilhas Excel (*.xlsx);;Todos os Arquivos (*)", options=options)
        # Adicione aqui a lógica para abrir a planilha de chamados
        pass

def main():
    app = QApplication(sys.argv)
    forms_flow_app = FormsFlowApp()
    forms_flow_app.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
