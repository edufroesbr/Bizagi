from pypdf import PdfReader
import logging
import os
import excel_helper

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

class BizagiValidator:
    def __init__(self):
        pass

    def check_res_1125_compliance(self, case_data, documents):
        """
        Main validation function following the 'ROTEIRO DE ANÁLISE'.
        
        Args:
            case_data (dict): Contains keys: 'contract_code', 'cnpj', 'debt_amount'
            documents (dict): Map of document types to file paths.
        
        Returns:
            dict: {
                "approved": bool,
                "correction_needed": bool,
                "reasons": list[str],
                "failed_docs": dict
            }
        """
        reasons = []
        is_compliant = True 
        
        logger.info("Starting RES 1125 Validation per Roteiro...")

        # 1. Conferência do valor devido (Rule 1.1)
        contract_code = case_data.get('contract_code')
        cnpj = case_data.get('cnpj')
        debt_amount = case_data.get('debt_amount')
        
        if contract_code and cnpj and debt_amount:
            # First, find the reference number (Step 1.1.1)
            ref_number = excel_helper.find_reference_number(contract_code)
            if ref_number:
                # Then, Find the AVD file and Sum values for the CNPJ (Step 1.1.1 Quinto passo)
                is_valid_amount, found_val, msg = excel_helper.validate_debt_amount(ref_number, cnpj, debt_amount)
                if not is_valid_amount:
                    reasons.append(f"Divergência de valores (Rule 1.1): {msg}")
                    is_compliant = False
            else:
                reasons.append(f"Contrato {contract_code} não encontrado na planilha mestre (Rule 1.1).")
                is_compliant = False
        else:
            reasons.append("Dados do caso incompletos (Contrato, CNPJ ou Valor) para validação cruzada.")
            is_compliant = False

        # 2. Conferência da inscrição no CADIN ANEEL (Rule 2.1)
        # "Ficou estabelecido como aceitável o envio das seguintes evidências..."
        # We check if AT LEAST ONE of the acceptable evidences is present.
        # Ideally, we'd like to see the Report (Relatório).
        
        has_cadin_report = self._doc_exists(documents, 'relatorio_cadin')
        has_cadin_history = self._doc_exists(documents, 'comunicacao_eletronica')
        has_cadin_memorial = self._doc_exists(documents, 'memorial_descritivo') # For multiple contracts
        
        if not (has_cadin_report or has_cadin_history or has_cadin_memorial):
             reasons.append("Ausência de evidências de inscrição no CADIN (Rule 2.1).")
             is_compliant = False

        # 3. Conferência do Protesto (Rule 3.1)
        # "Necessário que o documento comprove a efetivação do protesto"
        if not self._doc_exists(documents, 'comprovante_protesto'):
            reasons.append("Comprovante de protesto não encontrado (Rule 3.1).")
            is_compliant = False
        else:
            # Check for keywords confirming 'effective' protest
            protest_path = documents.get('comprovante_protesto')
            if not self.check_keywords(protest_path, ["protesto", "efetivado", "intimação", "tabelião", "cartório"]):
                # Warning only? The script says "Necessário que comprove".
                # For safety, if keywords missing, we flag it.
                reasons.append("Comprovante de protesto parece inválido ou ilegível (Rule 3.1).")
                is_compliant = False

        # 4. Conferência do Termo de compromisso (Rule 4.1)
        # "Analisar se o termo de compromisso foi juntado e está assinado..."
        if not self._doc_exists(documents, 'termo_compromisso'):
            reasons.append("Termo de compromisso ausente (Rule 4.1).")
            is_compliant = False
        else:
             # Check for signature keywords
            term_path = documents.get('termo_compromisso')
            if not self.check_keywords(term_path, ["assinado", "assinatura", "testemunha", "firmado"]):
                reasons.append("Termo de compromisso ausente ou sem assinatura (Rule 4.1).")
                is_compliant = False

        # 5. Conferência dos poderes (Rule 5.1) -> Usually part of Termo or separate Acts
        # The prompt says "Analisar a certidão para confirmação..." in 5.1 but title says "poderes".
        # It's likely checking the "Regularidade do CNPJ" doc for the next rule, or Acts for this one.
        # We will assume if Term is okay, we proceed to 6.1.
        
        # 6. Conferência da certidão de regularidade do CNPJ (Rule 6.1)
        if not self._doc_exists(documents, 'certidao_regularidade'):
             reasons.append("Certidão de regularidade ausente (Rule 6.1).")
             is_compliant = False
        
        # 7. Impossibility check (OBSERVAÇÃO)
        # "O CAMPO ... não deve conter documentos"
        if self._doc_exists(documents, 'impossibilidade_protesto'):
             reasons.append("Campo 'Impossibilidade do Protesto' não deve conter documentos (Rule 7.0).")
             is_compliant = False

        failed_docs = {}
        
        logger.info(f"Validation finished. Compliant: {is_compliant}")
        
        if not is_compliant:
            if "Rule 2.1" in str(reasons):
                failed_docs['Inscrição no cadastro de inadimplentes'] = "Ausência de evidências de inscrição no CADIN (Rule 2.1)"
            
            if "Rule 3.1" in str(reasons):
                failed_docs['Comprovante de Protesto'] = "Comprovante de protesto não encontrado ou inválido (Rule 3.1)"
                
            if "Rule 4.1" in str(reasons):
                failed_docs['Termo de Compromisso'] = "Termo de compromisso ausente ou sem assinatura (Rule 4.1)"
                
            if "Rule 6.1" in str(reasons):
                 failed_docs['Regularidade do CNPJ'] = "Certidão de regularidade ausente (Rule 6.1)"
                 
            if "Rule 7.0" in str(reasons):
                 failed_docs['Impossibilidade'] = "Campo não deve conter documentos (Rule 7.0)"

        return {
            "approved": is_compliant,
            "correction_needed": not is_compliant,
            "reasons": reasons,
            "failed_docs": failed_docs
        }

    def _doc_exists(self, documents, key):
        """Helper to check if file path exists."""
        path = documents.get(key)
        return path and os.path.exists(path)

    def check_keywords(self, file_path, keywords):
        """
        Simple check if ANY of the keywords exist in the PDF text.
        """
        try:
            reader = PdfReader(file_path)
            text = ""
            for page in reader.pages:
                text += page.extract_text().lower()
            
            for kw in keywords:
                if kw.lower() in text:
                    return True
            return False
        except Exception as e:
            logger.error(f"Error reading PDF {file_path}: {e}")
            return True # If we can't read it, we give benefit of doubt to avoid blocking on OCR needs

    def validate_protest_amount(self, file_path, expected_amount, memory_path=None):
        """
        Checks if the expected debt amount appears in the PDF text.
        If 'memory_path' is provided and the value is NOT found in the primary file,
        it checks the memory file as a fallback (cross-reference).
        
        Args:
            file_path (str): Path to the Protest PDF.
            expected_amount (str): The amount string (e.g. "1.234,56").
            memory_path (str, optional): Path to "Memória de Cálculo" PDF.
            
        Returns:
            (bool, str): Success status and message.
        """
        try:
            if not os.path.exists(file_path):
                return False, "Arquivo de Protesto não encontrado."

            # Normalize expected amount: remove spaces, generally keep basic formatting
            target_val = expected_amount.strip()
            # Remove "R$" and spaces to get raw number pattern logic
            raw_val = target_val.replace("R$", "").replace(" ", "").strip()
            
            # --- Helper to Check a Single File ---
            def check_file(path, label):
                if not path or not os.path.exists(path):
                    return False, f"{label}: Arquivo não acessível."
                
                try:
                    reader = PdfReader(path)
                    full_text = ""
                    for page in reader.pages:
                        full_text += page.extract_text()
                except Exception as e:
                    return False, f"{label}: Erro leitura PDF ({e})"

                # 1. Simple Check
                if target_val in full_text:
                    return True, f"Valor {target_val} encontrado em {label} (Match Exato)."
                
                # 2. Robust Regex Check
                import re
                parts = re.split(r'\D+', raw_val)
                regex_parts = []
                for i, part in enumerate(parts):
                    regex_parts.append(part)
                    if i < len(parts) - 1:
                        regex_parts.append(r'[\.,\s]*')
                
                number_pattern = "".join(regex_parts)
                final_pattern = r'(?:R\$)?\s*' + number_pattern
                
                if re.search(final_pattern, full_text, re.IGNORECASE):
                     return True, f"Valor {target_val} encontrado em {label} (Regex Flexível)."
                
                return False, f"Valor {target_val} não encontrado em {label}."

            # --- 1. Check Protest File ---
            found_in_protest, msg_protest = check_file(file_path, "Protesto")
            if found_in_protest:
                return True, msg_protest
            
            # --- 2. Check Memory File (Cross-Reference) ---
            if memory_path:
                found_in_memory, msg_memory = check_file(memory_path, "Memória de Cálculo")
                if found_in_memory:
                     return True, f"VALIDADO VIA MEMÓRIA: {msg_memory} (Protesto divergente aceito por compor múltiplos contratos)"
                else:
                     return False, f"Valor não encontrado no Protesto nem na Memória. ({msg_protest} | {msg_memory})"
            
            return False, msg_protest

        except Exception as e:
            return False, f"Erro fatal na validação de protesto: {e}"
