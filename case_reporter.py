import csv
import os
import datetime
import config

class CaseReporter:
    def __init__(self, report_filename="case_report.csv"):
        self.report_path = os.path.join(config.PROCESS_DIR, report_filename)
        self.headers = [
            "CaseID", "Timestamp_Start", "Contract_Code", "CNPJ", 
            "Debt_Amount", "Visual_Validation_Status", "Doc_Validation_Status", 
            "Final_Action", "Timestamp_End"
        ]
        self.current_case_data = {}
        self._initialize_csv()

    def _initialize_csv(self):
        """Creates the CSV file with headers if it doesn't exist."""
        if not os.path.exists(self.report_path):
            try:
                with open(self.report_path, mode='w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f, delimiter=';')
                    writer.writerow(self.headers)
                print(f"Arquivo de relatório criado em: {self.report_path}")
            except Exception as e:
                print(f"Erro ao criar arquivo de relatório: {e}")

    def start_case(self, case_id):
        """Initializes tracking for a new case."""
        self.current_case_data = {
            "CaseID": case_id,
            "Timestamp_Start": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "Contract_Code": "",
            "CNPJ": "",
            "Debt_Amount": "",
            "Visual_Validation_Status": "N/A",
            "Doc_Validation_Status": "N/A",
            "Final_Action": "Processing",
            "Timestamp_End": ""
        }
        print(f" Relatório: Iniciado rastreamento para o caso {case_id}")

    def update_data(self, key_map):
        """Updates specific fields in the current tracking session."""
        valid_keys = ["Contract_Code", "CNPJ", "Debt_Amount"]
        for key, value in key_map.items():
            # Map bot internal keys to CSV headers if needed, or pass directly
            # Here we assume key_map uses internal keys like 'contract_code' -> 'Contract_Code'
            if key == 'contract_code': self.current_case_data["Contract_Code"] = value
            elif key == 'cnpj': self.current_case_data["CNPJ"] = value
            elif key == 'debt_amount': self.current_case_data["Debt_Amount"] = value

    def log_visual_validation(self, status, message=""):
        """Logs the result of the visual Excel validation."""
        self.current_case_data["Visual_Validation_Status"] = f"{status} - {message}"

    def log_doc_validation(self, status, details=""):
        """Logs the result of the document analysis."""
        self.current_case_data["Doc_Validation_Status"] = f"{status} - {details}"

    def log_step(self, step_name, status):
        """Generic step logger (updates a dynamic or aggregated field if needed)."""
        # For simplicity, we might just print or append to a Steps column if we added one.
        # But based on current CSV structure, maybe we reuse Doc_Validation_Status or add logic.
        print(f" Relatório Step [{step_name}]: {status}")
        # Append to Doc Validation for tracking
        current = self.current_case_data.get("Doc_Validation_Status", "")
        self.current_case_data["Doc_Validation_Status"] = f"{current} | {step_name}: {status}"

    def finalize_case(self, action):
        """Writes the final case data to the CSV."""
        self.current_case_data["Final_Action"] = action
        self.current_case_data["Timestamp_End"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        try:
            with open(self.report_path, mode='a', newline='', encoding='utf-8') as f:
                writer = csv.DictWriter(f, fieldnames=self.headers, delimiter=';')
                writer.writerow(self.current_case_data)
            print(f" Relatório: Caso {self.current_case_data.get('CaseID')} salvo com sucesso.")
        except Exception as e:
            print(f"Erro ao salvar relatório do caso: {e}")
        
        # Reset current data
        self.current_case_data = {}
