"""Automation_CRM2"""

import time
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver import ActionChains
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
        # Se agrega la opcion para que use una instanacia de google exitente
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

        time.sleep(3)

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
        time.sleep(1)
        name_group.send_keys(group_name)

        # Coloca la descripcion del grupo
        descrip_group = self.driver.find_element(
            By.XPATH,
            "/html/body/div[3]/div[2]/div/mat-dialog-container/app-admin-groups/form/mat-dialog-content/mat-form-field[3]/div/div[1]/div/input",
        )
        time.sleep(1)
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

        # Varible boton de guardar grupo
        # save_group = self.driver.find_element(
        #     By.XPATH,
        #     "/html/body/div[2]/div[2]/div/mat-dialog-container/app-admin-groups/form/mat-dialog-actions/button[2]",
        # )
        # save_group.click()

    # Sección Inicio/Características

    def create_form(
        self,
        form_name,
        group_name,
        name_rol_list,
        option_yes_not_list,
        max_attempts=30,
        delay=1,
    ):
        """Metodo para colocar las caracteristicas del formulario"""

        # Boton ingreso a modulo Fomularios
        button_forms = self.wait.until(
            EC.visibility_of_element_located(
                (
                    By.XPATH,
                    "/html/body/app-root/app-mios/app-side-bar/div/mat-sidenav-container/mat-sidenav/div/app-left-nav/div/div/div/mat-nav-list/div/a[1]",
                )
            )
        )

        button_forms.click()

        # Boton Crear Formulario
        create_button = self.wait.until(
            EC.visibility_of_element_located(
                (
                    By.XPATH,
                    "/html/body/app-root/app-mios/app-side-bar/div/mat-sidenav-container/mat-sidenav-content/div/app-forms-list/div/div/div[2]/div/button",
                )
            )
        )
        create_button.click()

        # Coloca el nombre del Formulario
        name_form = self.wait.until(
            EC.visibility_of_element_located(
                (
                    By.XPATH,
                    "/html/body/app-root/app-mios/app-side-bar/div/mat-sidenav-container/mat-sidenav-content/div/app-admin-forms/div/div[1]/div[2]/mat-form-field[1]/div/div[1]/div/input",
                )
            )
        )
        name_form.send_keys(form_name)

        # Elige el tipo del formulario
        type_form = self.driver.find_element(
            By.XPATH,
            "/html/body/app-root/app-mios/app-side-bar/div/mat-sidenav-container/mat-sidenav-content/div/app-admin-forms/div/div[1]/div[2]/mat-form-field[2]/div/div[1]/div/mat-select",
        )
        type_form.click()

        option_type = self.wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "/html/body/div[3]/div[2]/div/div/div/mat-option[3]")
            )
        )
        option_type.click()

        # Elige que roles descargan
        download_general = self.driver.find_element(
            By.XPATH,
            "/html/body/app-root/app-mios/app-side-bar/div/mat-sidenav-container/mat-sidenav-content/div/app-admin-forms/div/div[1]/div[2]/mat-form-field[3]/div/div[1]/div/mat-select",
        )
        download_general.click()

        # Bucle para Elegir los roles
        for name_rol in name_rol_list:
            for i in range(4):
                try:
                    option = self.wait.until(
                        EC.visibility_of_element_located(
                            (By.XPATH, f"//mat-option[contains(., '{name_rol}')]")
                        )
                    )
                    break
                except (TimeoutException, NoSuchElementException):
                    time.sleep(delay)
            if option:
                option.click()
            else:
                raise ValueError(f"No se puede encontrar el rol {name_rol}")

        time.sleep(2)

        download_general.send_keys(Keys.TAB)

        # Elige la campaña a la que pertenece
        campaigns = self.driver.find_element(
            By.XPATH,
            "/html/body/app-root/app-mios/app-side-bar/div/mat-sidenav-container/mat-sidenav-content/div/app-admin-forms/div/div[1]/div[2]/mat-form-field[4]/div/div[1]/div/mat-select",
        )

        campaigns.click()

        select_campaigns = self.wait.until(
            EC.visibility_of_element_located(
                (By.XPATH, "/html/body/div[3]/div[2]/div/div/div/mat-option[42]")
            )
        )
        select_campaigns.click()
        time.sleep(2)

        # Elige el grupo
        group = self.driver.find_element(
            By.XPATH,
            "/html/body/app-root/app-mios/app-side-bar/div/mat-sidenav-container/mat-sidenav-content/div/app-admin-forms/div/div[1]/div[2]/mat-form-field[5]/div/div[1]",
        )

        group.click()

        # Bucle para buscar el grupo por nombre
        for attempt in range(max_attempts):
            try:
                group_option = self.wait.until(
                    EC.element_to_be_clickable(
                        (By.XPATH, f"//mat-option[contains(., '{group_name}')]")
                    )
                )
                break
            except (TimeoutException, NoSuchElementException):
                time.sleep(delay)
        if group_option:
            group_option.click()
        else:
            raise ValueError("No se puede encontrar el grupo")

        time.sleep(2)

        # Bucle para seleccionar opciones si y no los dos campos siguientes
        for i, option_yes_not in enumerate(option_yes_not_list):

            # Elemento desplegable
            time_of_tipifi = self.driver.find_element(
                By.XPATH,
                f"/html/body/app-root/app-mios/app-side-bar/div/mat-sidenav-container/mat-sidenav-content/div/app-admin-forms/div/div[1]/div[2]/mat-form-field[{i+6}]",
            )

            time_of_tipifi.click()
            time.sleep(2)
            if option_yes_not == "si":
                # Opciones si y no
                option_tipifi1 = self.wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "/html/body/div[3]/div[2]/div/div/div/mat-option[1]",
                        )
                    )
                )
                option_tipifi1.click()
            else:
                option_tipifi2 = self.wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "/html/body/div[3]/div[2]/div/div/div/mat-option[2]",
                        )
                    )
                )
                option_tipifi2.click()

    def action_create(
        self,
        type_camp,
        name_campo,
        num_colum,
        list_yes_no,
        list_yes_no2,
        name_rol_see_list,
        name_rol_edit_list,
        chacter_min,
        chater_max,
        place_num,
    ):
        """Metodo Para iniciar la creación del formulario"""

        # Tiempo para ubicar el campo
        time.sleep(0.5)

        # Diccionario con los campos y su valor
        campos = {
            "texto": 1,
            "desplegable": 2,
            "multipleseleccion": 3,
            "fecha": 4,
            "agendamiento": 5,
            "numerico": 6,
            "comentario": 7,
            "email": 8,
            "radiobutton": 9,
            "autocomplete": 10,
            "moneda": 11,
            "archivo": 12,
            "tiempo": 13,
        }

        # Almacena el valor por nombre de campo
        num_campo = campos.get(type_camp.lower())

        # Manejo de error en caso de no encotrar el nombre del campo
        if num_campo is None:
            print(f"Tipo de campo no reconocido: {type_camp}")
            print("Tipos de campo válidos:", list(campos.keys()))
            return

        # Base del xpath de los campos
        base_camp_xpath = "/html/body/app-root/app-mios/app-side-bar/div/mat-sidenav-container/mat-sidenav-content/div/app-admin-forms/div/div[2]/mat-tab-group/div/mat-tab-body[1]/div/div/div/div/div"

        # xpath completo del campo
        camp_xpath = f"{base_camp_xpath}[{num_campo}]"

        # Base del xpath del placeholder
        base_xpath = "/html/body/app-root/app-mios/app-side-bar/div/mat-sidenav-container/mat-sidenav-content/div/app-admin-forms/div/div[1]/div[3]/mat-card/mat-grid-list/div/mat-grid-tile"

        # Estructura del xpath dell placeholder
        if place_num == 1:
            placeholder_xpath = f"{base_xpath}/figure/div"
        else:
            placeholder_xpath = f"{base_xpath}[{place_num}]/figure/div"

        # Elementos campo y placeholder con las modificaciones realizadas
        campo = self.driver.find_element(By.XPATH, camp_xpath)
        placeholder = self.driver.find_element(By.XPATH, placeholder_xpath)

        # Secuencia de acciones para el movimiento del campo hacia el placeholder
        action = ActionChains(self.driver)
        action.move_to_element(campo)
        action.click_and_hold()
        action.pause(0.5)
        action.move_to_element(placeholder)
        action.move_by_offset(5, 5)
        action.release()
        action.perform()

        # Tiempo para cargar la parte de variables
        time.sleep(1)

        # Imprime las variables de tiene cada campo para verificar que se esten obteniendo de manera exitosa
        print(name_campo)
        print(num_colum)
        for option in list_yes_no:
            print(option)
        for option2 in list_yes_no2:
            print(option2)
        for rol in name_rol_edit_list:
            print(rol)
        for rol2 in name_rol_see_list:
            print(rol2)
        print(chacter_min)
        print(chater_max)

    def process_excel(self, excel_path):
        """Metodo para procesar los datos del excel"""

        campos_agregar = load_workbook(excel_path)
        excel = campos_agregar.active

        # Bucle para tomar los datos las celdas
        for index, row in enumerate(
            excel.iter_rows(min_row=2, values_only=True), start=1
        ):
            # Varibles con las caracteristicas de los campos
            type_camp = row[0]
            name_campo = row[1]
            num_colum = str(row[2])
            list_yes_no = [cell for cell in row[3:8] if cell in ["si", "no"]]
            list_yes_no2 = [cell for cell in row[8:15] if cell in ["si", "no"]]
            name_rol_see_list = str(row[15]).split(",") if row[15] else []
            name_rol_edit_list = str(row[16]).split(",") if row[16] else []
            chacter_min = str(row[17])
            chater_max = str(row[18])

            # Ejecución del método action que toma las variables extraidas del excel
            self.action_create(
                type_camp,
                name_campo,
                num_colum,
                list_yes_no,
                list_yes_no2,
                name_rol_see_list,
                name_rol_edit_list,
                chacter_min,
                chater_max,
                index,
            )


def main():
    """Metodo para ejecutar los metodos definidos"""

    try:
        # Instancia de la clase
        Automation = CRM2Automation()

        # Ejecución método sección Usuarios por Excel
        users_to_add = Automation.read_user_from_excel(
            r"C:\Users\USUARIO\Downloads\Campos.xlsx", "usuarios", "A2", "A3"
        )

        # Ejecución método sección Inicio/Creación
        Automation.create_group("AXA AUTOS", "AXA AUTOS", users_to_add)

        # Ejecución método sección Inicio/Características
        Automation.create_form(
            "Formulario Express AXA",
            "AXA AUTOS",
            ["Administrador", "Supervisor CRM", "Asesor CRM", "BackOffice"],
            ["si", "no"],
        )

        # Ejecución método sección Logica Integración Excel
        Automation.process_excel(r"C:\Users\USUARIO\Downloads\Campos.xlsx")

    except ImportError as e:
        print(f"SE PRODUJO UN ERROR EN: {e}")


if __name__ == "__main__":
    main()
