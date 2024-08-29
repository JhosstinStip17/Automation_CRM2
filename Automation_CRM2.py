"""Automation_CRM2"""

import time
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException


class CRM2Automation:
    """Clase Automatización"""

    def __init__(self):
        """Metodo contructor de la clase"""

        # Se crea las opciones para abrir una instancia de google existente
        self.chrome_option = Options()
        self.chrome_option.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        # Se crea el driver
        self.driver = webdriver.Chrome(options=self.chrome_option)
        # Se define el tiempo de espera para el driver
        self.wait = WebDriverWait(self.driver, 10)

    def read_user_from_excel(self, file_path, sheet_name, star_cell, end_cell):
        """Metodo para leer los usuarios de un archivo Excel"""

        # Variable para alamcenar el documento excel
        excel = load_workbook(filename=file_path)
        # Entrar en al hoja del excel
        sheet = excel[sheet_name]
        # Lista para alamcenar los usuarios
        users = []
        # Bucle para buscar a los usuarios y almacenarlos en la lista
        for row in sheet[star_cell:end_cell]:
            for cell in row:
                if cell.value is not None:
                    # Se coloca str para converti el valor en un string
                    users.append(str(cell.value))
        return users

    def create_group(self, group_name, group_descrip, users_to_add):
        """Metodo para crear el grupo"""

        # Varible modulo grupo
        group_button = self.driver.find_element(
            By.XPATH,
            "/html/body/app-root/app-mios/app-side-bar/div/mat-sidenav-container/mat-sidenav/div/app-left-nav/div/div/div/mat-nav-list/div/a[3]",
        )
        group_button.click()

        time.sleep(2)

        # Variable al boton crear grupo
        create_group = self.wait.until(
            EC.visibility_of_element_located(
                (
                    By.XPATH,
                    "/html/body/app-root/app-mios/app-side-bar/div/mat-sidenav-container/mat-sidenav-content/div/app-groups-list/div/div/div[1]/div[2]/div/button",
                )
            )
        )
        create_group.click()

        time.sleep(2)

        # Variable a select de campañas dentro de las caracteristicas de crear grupo
        campaigns = self.wait.until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "/html/body/div[3]/div[2]/div/mat-dialog-container/app-admin-groups/form/mat-dialog-content/mat-form-field[1]",
                )
            )
        )

        campaigns.click()
        time.sleep(2)

        # Seleciona el valor dentro del select
        select_campaign = self.wait.until(
            EC.visibility_of_element_located(
                (By.XPATH, "/html/body/div[3]/div[4]/div/div/div/mat-option[42]")
            )
        )
        select_campaign.click()

        time.sleep(5)

        # Sale del select
        out_select = self.driver.find_element(By.XPATH, "/html/body/div[3]/div[3]")
        out_select.click()

        time.sleep(2)

        # Coloca el nombre del grupo
        name_group = self.wait.until(
            EC.visibility_of_element_located(
                (
                    By.XPATH,
                    "/html/body/div[3]/div[2]/div/mat-dialog-container/app-admin-groups/form/mat-dialog-content/mat-form-field[2]/div/div[1]/div/input",
                )
            )
        )
        time.sleep(2)
        name_group.send_keys(group_name)

        # Coloca la descripcion del grupo
        descrip_group = self.driver.find_element(
            By.XPATH,
            "/html/body/div[3]/div[2]/div/mat-dialog-container/app-admin-groups/form/mat-dialog-content/mat-form-field[3]/div/div[1]/div/input",
        )
        time.sleep(2)
        descrip_group.send_keys(group_descrip)

        # Bucle para añadir a los usuarios
        for user_to_add in users_to_add:

            add_user = self.driver.find_element(
                By.XPATH,
                "/html/body/div[3]/div[2]/div/mat-dialog-container/app-admin-groups/form/mat-dialog-content/div[4]/div[1]/mat-form-field/div/div[1]/div[1]/input",
            )
            add_user.send_keys(user_to_add)
            add_user.send_keys(Keys.ENTER)
            time.sleep(2)

            button_add = self.wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "/html/body/div[3]/div[2]/div/mat-dialog-container/app-admin-groups/form/mat-dialog-content/div[4]/div[2]/button",
                    )
                )
            )
            button_add.click()

        time.sleep(4)

        cancelar = self.driver.find_element(
            By.XPATH,
            "/html/body/div[3]/div[2]/div/mat-dialog-container/app-admin-groups/form/mat-dialog-actions/button[1]",
        )

        cancelar.click()
        time.sleep(2)

        # # Varible boton de guardar grupo
        # save_group = self.driver.find_element(
        #     By.XPATH,
        #     "/html/body/div[2]/div[2]/div/mat-dialog-container/app-admin-groups/form/mat-dialog-actions/button[2]",
        # )
        # save_group.click()


def main():
    """Metodo para ejecutar los metodos definidos"""

    try:
        Automation = CRM2Automation()
        users_to_add = Automation.read_user_from_excel(
            r"C:\Users\USUARIO\Downloads\usuarios.xlsx", "usuarios", "A2", "A3"
        )
        Automation.create_group("Nombre Grupos", "Descripcion del grupo", users_to_add)
    except ImportError as e:
        print(f"SE PRODUJO UN ERROR EN: {e}")


if __name__ == "__main__":
    main()
