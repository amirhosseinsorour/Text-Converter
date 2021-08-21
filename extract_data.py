class PtuWorkBook:
    """
    This class contains all of data in a ptu file in form of subclasses
    Attributes will be used to show in excel file, or convert to another format
    """

    def __init__(self):
        self.name = ""
        self.preface = self.Preface()
        self.include = []
        self.comment = []
        self.user_code = []  # Array of UserCode objects
        self.stub_definitions = []  # Array of DefineStub objects
        self.initialisation = []  # Array of Initialisation objects
        self.environments = []  # Array of Environment objects
        self.services = []  # Array of Service objects

    class Preface:
        """
        This subclass contains initial data about ptu file
        """

        def __init__(self):
            self.purpose = ""
            self.processor = ""
            self.tool_chain = ""
            self.header = self.Header()

        class Header:
            """
            This subclass contains important datarmation mentioned in header file.
            This datarmation is used in Preface(superclass) as an attribute.
            """

            def __init__(self):
                self.module_name = ""
                self.module_version = ""
                self.test_plan_version = ""

    class UserCode:
        """
        UserCode can appear anywhere in ptu file
        So this subclass is used in PtuWorkBook(superclass) , Service and Element class for better grouping
        All UserCode lines start with '#'
        """

        def __init__(self):
            self.conditions = []
            self.code = ""

    class Stub:
        """
        STUB instruction describes all calls to a simulated function in a test script.
        This class is used in ELEMENT and DefineStub classes.
        """

        def __init__(self):
            self.conditions = []
            self.stub_definition = ""

    class DefineStub:
        """
        DEFINE STUB instructions are optional and has a scope in format: "DEFINE STUB ... END DEFINE"
        The DEFINE STUB and END DEFINE instructions delimit a simulation block
        consisting of stub definition functions, methods, or procedure declarations.
        """

        def __init__(self):
            self.name = ""
            self.stub_list = []  # Array of Stub objects
            self.conditions = []

    class Initialisation:
        """
        Initialisation in ptu files is optional.
        Initialisation has a scope in format: "INITIALISATION ... END INITIALISATION"
        Content of initialisation is saved in "description" attribute.
        """

        def __init__(self):
            self.conditions = []
            self.description = ""

    class TestCase:
        """
        TestCases appear in ENVIRONMENT and ELEMENT classes.
        In every test case, there is a parameter that is being tested.
        This parameter has type, name, initial and expected value.
        """

        def __init__(self):
            self.conditions = []
            self.param_type = ""
            self.param_name = ""
            self.init = ""
            self.ev = ""

    class Environment:
        """
        ENVIRONMENT instruction defines a test environment declaration in ptu files.
        ENVIRONMENT has a scope in format: "ENVIRONMENT ... END ENVIRONMENT"
        All of test cases are saved in an array of TestCase objects, named "test_case_list"
        """

        def __init__(self):
            self.name = ""
            self.test_case_list = []  # Array of TestCase objects
            self.conditions = []

    class Service:
        """
        Services are functions that are being tested.
        Every Service consists of several tests and every test consist of an element.
        These tests are saved in an array, named "test_list", in format of TEST objects.
        Every service may use several environments, which are saved in "use" array.
        """

        def __init__(self):
            self.name = ""
            self.test_list = []  # Array of Test objects
            self.conditions = []
            self.user_code = []  # Array of UserCode objects
            self.use = []  # Array of strings (USEs)

        def has_user_code(self):
            """
            :return: True if Service has any user-code(itself or its tests), False if not
            """
            if self.user_code:
                return True
            for test in self.test_list:
                if test.element.user_code:
                    return True
            return False

        def extend_use(self):
            """
            Every USE in a service belongs to all of elements in its sub-tests.
            This function adds USEs of service to all of the elements in its sub-tests.
            """
            for test in self.test_list:
                element = test.element
                element.use.extend(self.use)

        def get_all_user_code(self):
            """
            :return: All of the user-code of this service in form of one string.
            """
            all_user_code = [u.code for u in self.user_code]
            return "\n".join(all_user_code)

    class Element:
        """
        Every Test of every service, consists of one element.
        All of parameters that are going to be tested in a particiular TEST, are in element in 3 formats:
        Input-data, Output-data, Calibrations. All of these are objects of TestCase Class.
        """

        def __init__(self):
            self.use = []
            self.input_data = []  # Array of TestCase objects
            self.calibrations = []  # Array of TestCase objects
            self.output_data = []  # Array of TestCase objects
            self.stub_calls = []
            self.user_code = []  # Array of strings (USEs)

        def get_all_data(self):
            """
            :return: All of test-cases including input-data, output-data, and calibrations.
            """
            return self.input_data + self.calibrations + self.output_data

    class Test:
        """
        Every Service consists of several tests.
        There is a single element in every test.
        """

        def __init__(self):
            self.name = ""
            self.family = ""
            self.comment = []
            self.use = []
            self.element = PtuWorkBook.Element()
            self.conditions = []

        def extend_use(self):
            """
            Every USE in a test belongs to all of its sub-elements.
            This function adds USEs of test to all of the elements in its sub-elements.
            """
            self.element.use.extend(self.use)


class LineNum:
    """
    This class is used for saving line numbers of specific data in PTU files.
    Line numbers will be used in extracting data and saving them in an object of PtuWorkbook class./
    """

    class ServiceLineNum:
        """
        For sub-elements of SERVICE, we define a new class because we have a list of services in a PTU file.
        So we have a list of ServiceLineNum in LineNum class and a list of Service in PtuWorkBook class.
        """

        def __init__(self):
            self.START = 0
            self.TEST_START = []
            self.ELEMENT_START = []
            self.ELEMENT_END = []
            self.TEST_END = []
            self.END = 0

        def get_test_count(self):
            """
            :return: Number of Tests in PTU file.
            """
            return len(self.TEST_START)

    class IfScope:
        """
        For IF-scopes, we save them here in this class.
        We save Start and End Line num of this scope and also its condition.
        """

        def __init__(self):
            self.START_IF = 0
            self.END_IF = 0
            self.condition = ""

    def __init__(self):
        self.PURPOSE = 0
        self.PROCESSOR = 0
        self.TOOL_CHAIN = 0
        self.INCLUDE_LIST = []
        self.HEADER_START = 0
        self.COMMENT_START = 0
        self.COMMENT_END = 0
        self.USER_CODE_LIST = []
        self.DEFINE_STUB_START_LIST = []
        self.DEFINE_STUB_END_LIST = []
        self.TEST_CASES_START_FLAG = False
        self.INITIALISATION_START = 0
        self.INITIALISATION_END = 0
        self.ENVIRONMENT_START_LIST = []
        self.ENVIRONMENT_END_LIST = []
        self.SERVICE_LIST = []  # Array of ServiceLineNum
        self.IF_SCOPE_LIST = []  # Array of IfScope

    def service_start(self, line_num):
        """
        Creating new object of ServiceLineNum class and add it to the list.
        :param line_num: Start line number of the Service scope
        """
        new_service = self.ServiceLineNum()
        new_service.START = line_num
        self.SERVICE_LIST.append(new_service)

    def test_start(self, line_num):
        """
        Adding start of a new Test to the last started Service.
        :param line_num: Start line number of the Test scope
        """
        current_service = self.SERVICE_LIST[-1]
        current_service.TEST_START.append(line_num)

    def test_end(self, line_num):
        """
        Adding end of the last started Test to the last started service
        :param line_num: End line number of the Test scope
        """
        current_service = self.SERVICE_LIST[-1]
        current_service.TEST_END.append(line_num)

    def element_start(self, line_num):
        """
        Adding start of a new Element to the last started Service.
        :param line_num: Start line number of the Element scope
        """
        current_service = self.SERVICE_LIST[-1]
        current_service.ELEMENT_START.append(line_num)

    def element_end(self, line_num):
        """
        Adding end of the last started Element to the last started service
        :param line_num: End line number of the Element scope
        """
        current_service = self.SERVICE_LIST[-1]
        current_service.ELEMENT_END.append(line_num)

    def service_end(self, line_num):
        """
        Adding End line number of the last started Service
        :param line_num: End line number of the Service scope
        """
        current_service = self.SERVICE_LIST[-1]
        current_service.END = line_num

    def get_define_stub_count(self):
        return len(self.DEFINE_STUB_START_LIST)

    def get_environment_count(self):
        return len(self.ENVIRONMENT_START_LIST)

    def valid_data(self, line_num):
        """
        This function checks whether the line is user-code or not
        """
        # If we have reached start of the first Service
        if self.TEST_CASES_START_FLAG:
            return False
        # If we are in a DefineStub scope
        if len(self.DEFINE_STUB_START_LIST) > len(self.DEFINE_STUB_END_LIST):
            return False
        for counter in range(self.get_define_stub_count()):
            if self.DEFINE_STUB_START_LIST[counter] < line_num < self.DEFINE_STUB_END_LIST[counter]:
                return False
        return True


def open_ptu(path, file_name):
    """
    This function opens PTU file and calls other functions to extract data.
    :returns: Object of PtuWorkBook class with all of data saved in it.
    """
    ptu_file = open(path, "rt")
    ptu_workbook = PtuWorkBook()
    ptu_workbook.name += file_name.replace(".ptu", "")
    ptu_file_lines, line_numbers = pre_process(ptu_file)
    classify_data(ptu_workbook, ptu_file_lines, line_numbers)

    return ptu_workbook


def save_line_num_by_type(line, line_numbers, number):
    """
    According to the type of keyword in the line, this function saves scopes of diffrent types of commands.
    """
    if line.upper().startswith("SERVICE "):  # If the line starts a SERVICE scope
        line_numbers.service_start(number)
    elif line.upper().startswith("END SERVICE"):  # If the line ends a SERVICE scope
        line_numbers.service_end(number)
    elif line.upper().startswith("TEST"):  # If the line starts a TEST scope
        line_numbers.test_start(number)
    elif line.upper().startswith("END TEST"):  # If the line ends a TEST scope
        line_numbers.test_end(number)
    elif line.upper().startswith("ELEMENT"):  # If the line starts a ELEMENT scope
        line_numbers.element_start(number)
    elif line.upper().startswith("END ELEMENT"):  # If the line ends a ELEMENT scope
        line_numbers.element_end(number)
    elif line.upper().startswith("INITIALISATION"):  # If the line starts a INITIALISATION scope
        line_numbers.INITIALISATION_START = number
    elif line.upper().startswith("END INITIALISATION"):  # If the line ends a INITIALISATION scope
        line_numbers.INITIALISATION_END = number
    elif line.upper().startswith("ENVIRONMENT"):  # If the line starts a ENVIRONMENT scope
        line_numbers.ENVIRONMENT_START_LIST.append(number)
    elif line.upper().startswith("END ENVIRONMENT"):  # If the line ends a ENVIRONMENT scope
        line_numbers.ENVIRONMENT_END_LIST.append(number)
    elif line.upper().startswith("DEFINE STUB"):  # If the line starts a DEFINE STUB scope
        line_numbers.DEFINE_STUB_START_LIST.append(number)
    elif line.upper().startswith("END DEFINE"):  # If the line ends a DEFINE STUB scope
        line_numbers.DEFINE_STUB_END_LIST.append(number)
    elif line.upper().startswith("COMMENT START"):  # If the line starts a COMMENT scope
        line_numbers.COMMENT_START = number
    elif line.upper().startswith("COMMENT END"):  # If the line ends a COMMENT scope
        line_numbers.COMMENT_END = number
    elif line.upper().startswith("HEADER"):
        line_numbers.HEADER_START = number
    elif line.startswith("##include"):
        line_numbers.INCLUDE_LIST.append(number)
    elif line.startswith("-- Purpose"):
        line_numbers.PURPOSE = number
    elif line.startswith("-- Processor"):
        line_numbers.PROCESSOR = number
    elif line.startswith("-- Tool chain"):
        line_numbers.TOOL_CHAIN = number
    elif line.startswith("-- Test Cases"):  # If the line indicates beginning of SERVICE scopes
        line_numbers.TEST_CASES_START_FLAG = True
    elif line.startswith("#") and line_numbers.valid_data(number):  # If the line includes user-code
        line_numbers.USER_CODE_LIST.append(number)


def save_if_scope(if_stack, lines, line_numbers, number):
    """
    This functions particularly saves scopes of IF and ELSEs.
    :param if_stack: LIFO list for saving scopes of IF.
    :param lines: All lines of PTU file
    :param line_numbers: Object of LineNum class
    :param number: Line number of IF,ELSE, or ENDIF command
    """
    line = lines[number].upper().strip()

    if line.startswith("IF"):
        """
        When we reach IF command, we save the line number as a start of an IF-scope.
        Then we save its condition and append it to the stack.
        """
        if_scope = LineNum.IfScope()
        if_scope.START_IF = number
        if_scope.condition += line.replace("IF", "").strip()
        if_stack.append(if_scope)
    elif line.startswith("ENDIF") or line.startswith("END IF"):
        """
        When we reach ENDIF command, we save the line number as a end of an IF-scope.
        Then we append the last IF-scope that we created into the list of all IF-scopes.
        """
        last_if = if_stack.pop()
        last_if.END_IF = number
        line_numbers.IF_SCOPE_LIST.append(last_if)
    elif line.startswith("ELSE"):
        """
        When we reach ELSE command, we save the line number as a end of the last IF-scope.
        Then we create new IF-scope with the the negated condition of the last IF-scope.
        """
        last_if = if_stack.pop()
        last_if.END_IF = number
        line_numbers.IF_SCOPE_LIST.append(last_if)
        line = lines[last_if.START_IF].upper().replace("IF", "IF NOT")
        lines[number] = line
        else_scope = LineNum.IfScope()
        else_scope.START_IF = number
        else_scope.condition += line.replace("IF NOT", "!").strip()
        if_stack.append(else_scope)


def pre_process(file_handler):
    """
    The pre-process part is going to save line numbers of all kind of data.
    This line numbers will be used to extract data.
    :return: ""lines"" as an array of strings, ""line_numbers as an object of LineNum class.
    """
    lines = file_handler.readlines()
    lines[0] = ""  # Making an empty element in lines array as an initial value
    file_handler.close()

    line_numbers = LineNum()
    if_stack = []

    for num in range(len(lines)):
        line = lines[num].strip()
        if line.find("--") > 0:  # If there is comment inside the line(not at the start of the line)
            lines[num] = line[:line.find("--")].strip()  # Removing comment inside the line
        if line.upper().startswith("IF") or line.upper().startswith("ELSE") or \
                line.upper().startswith("ENDIF") or line.upper().startswith("END IF"):
            save_if_scope(if_stack, lines, line_numbers, num)  # Saving IF-scope
        else:
            save_line_num_by_type(line, line_numbers, num)  # Saving all scopes (except IF) according to their type

    # Sorting all of IF-scopes by their start line number
    line_numbers.IF_SCOPE_LIST.sort(key=lambda i: i.START_IF, reverse=False)

    return lines, line_numbers


def classify_data(workbook, lines, line_numbers):
    """
    This function is the main function for extracting data from PTU file.
    It consists of several functions that each one is used for extracting different type of data.
    :param workbook: Object of PtuWorkBook class, for saving all data
    :param lines: Array of all af the lines in PTU file.
    :param line_numbers: Object of LineNum class, which helps us to find and extract data easier.
    """

    def extract_preface_data():
        """
        Extracting Preface data such as Purpose, Processor, Toolchain, and header info.
        """
        purpose = lines[line_numbers.PURPOSE]
        purpose = purpose.replace("-- Purpose:", "")
        workbook.preface.purpose += purpose.strip()

        processor = lines[line_numbers.PROCESSOR]
        processor = processor.replace("-- Processor:", "")
        workbook.preface.processor += processor.strip()

        tool_chain = lines[line_numbers.TOOL_CHAIN]
        tool_chain = tool_chain.replace("-- Tool chain:", "")
        workbook.preface.tool_chain += tool_chain.strip()

        header = lines[line_numbers.HEADER_START]
        header = header.replace("HEADER ", "")
        header = header.strip()
        # Header includes module name, module version, and test plan version that are seprated using ','
        try:
            header_module_name, header_module_version, header_test_plan_version = header.split(",")
            workbook.preface.header.module_name = header_module_name.strip()
            workbook.preface.header.module_version = header_module_version.strip()
            workbook.preface.header.test_plan_version = header_test_plan_version.strip()
        except ValueError:
            pass

    def extract_include_data():
        """
        Extracting list of included files from PTU file
        """
        for include_line_num in line_numbers.INCLUDE_LIST:
            include = lines[include_line_num]
            include = include.replace("##include ", "")
            workbook.include.append(include.strip())

    def extract_comment_data():
        """
        Extracting list of common comments by analyzing comment scope in PTU file
        """
        for line_num in range(line_numbers.COMMENT_START + 1, line_numbers.COMMENT_END):
            line = lines[line_num]
            if line.startswith("COMMENT "):
                line = line.replace("COMMENT ", "")
                workbook.comment.append(line.strip())

    def check_for_conditions(start_line_num, end_line_num):
        """
        A useful function for checking conditions of a given scope.
        This function checks scope of which conditions(IF-scopes) includes the scope which starts at start_line_num and ends at end_line_num
        :returns list of all IF-scopes with terms that we said in previous sentence.
        """
        condition_list = []
        for if_scope in line_numbers.IF_SCOPE_LIST:
            if if_scope.START_IF > end_line_num:  # If there is no IF-scope that includes the input scope
                break
            if start_line_num > if_scope.START_IF and end_line_num < if_scope.END_IF:
                condition_list.append(if_scope.condition)
        return condition_list

    def extract_user_code_data():
        """
        Extracting common user-code data in PTU file and grouping them by their conditions
        """
        for line_num_counter in range(len(line_numbers.USER_CODE_LIST)):
            line_num = line_numbers.USER_CODE_LIST[line_num_counter]
            if line_num == 0: continue  # If the line is a merged user-code
            line = lines[line_num].strip()
            user_code = PtuWorkBook.UserCode()
            user_code.code = line[1:]  # Remove '#'
            user_code.conditions = check_for_conditions(line_num, line_num)

            # Until the next user-code has same conditions as current line, we add it to current user-code
            while True:
                try:
                    next_line_num = line_numbers.USER_CODE_LIST[line_num_counter + 1]
                except IndexError:  # If the line was the last user-code line and has no next line
                    break
                if user_code.conditions != check_for_conditions(next_line_num, next_line_num):
                    break  # If next user-code doesn't have same conditions, don't merge it with current line

                # Merge next user-code with current user-code and check these terms for the following user-codes,too
                line_numbers.USER_CODE_LIST[line_num_counter + 1] = 0
                next_line = lines[next_line_num].strip()
                user_code.code += "\n" + next_line[1:]

                line_num_counter += 1

            workbook.user_code.append(user_code)

    def extract_stub_definitions_data():
        for counter in range(line_numbers.get_define_stub_count()):
            define_stub_name = lines[line_numbers.DEFINE_STUB_START_LIST[counter]]
            define_stub_name = define_stub_name.replace("DEFINE STUB ", "")
            define_stub = PtuWorkBook.DefineStub()
            define_stub.name = define_stub_name.strip()
            define_stub.conditions = check_for_conditions(line_numbers.DEFINE_STUB_START_LIST[counter],
                                                          line_numbers.DEFINE_STUB_END_LIST[counter])
            in_stub_scope = False
            tmp_stub_scope_start_line = 0
            tmp_line = ""
            for line_num in range(line_numbers.DEFINE_STUB_START_LIST[counter] + 1,
                                  line_numbers.DEFINE_STUB_END_LIST[counter]):
                line = lines[line_num].strip()
                if not line.startswith("#"):
                    continue
                line = line.replace("#", "")
                line = line.strip()
                if line.endswith(";") and not in_stub_scope:
                    stub = PtuWorkBook.Stub()
                    stub.conditions = check_for_conditions(line_num, line_num)
                    line = line.replace(";", "")
                    stub.stub_definition = line
                    define_stub.stub_list.append(stub)
                elif in_stub_scope:
                    tmp_line += "\n" + line
                    if "}" in line:
                        in_stub_scope = False
                        stub = PtuWorkBook.Stub()
                        stub.stub_definition = tmp_line
                        stub.conditions = check_for_conditions(tmp_stub_scope_start_line, line_num)
                        define_stub.stub_list.append(stub)
                        tmp_line = ""
                        tmp_stub_scope_start_line = 0
                elif not line.endswith(";") and not in_stub_scope:
                    tmp_line += line
                    tmp_stub_scope_start_line = line_num
                    in_stub_scope = True
            workbook.stub_definitions.append(define_stub)

    def extract_initialization_data():
        """
        Extracting data in Initialisation scope (If exists) and grouping them by their condition
        """
        for line_num in range(line_numbers.INITIALISATION_START + 1, line_numbers.INITIALISATION_END):
            line = lines[line_num].strip()

            # If line is empty or IF-scope (not in user-code format)
            if not line.startswith("#"): continue

            initialisation = PtuWorkBook.Initialisation()
            initialisation.description += line[1:]
            initialisation.conditions = check_for_conditions(line_num, line_num)

            # Checking whether next lines have the same list of conditions(IF-scope) or not
            while line_num < line_numbers.INITIALISATION_END:
                next_line_num = line_num + 1
                next_line = lines[next_line_num].strip()

                # If line is empty or IF-scope (not in user-code format)
                if not next_line.startswith("#"):
                    line_num += 1
                    continue

                if initialisation.conditions != check_for_conditions(next_line_num, next_line_num):
                    break

                # Merge next lines with current line and check these terms for the following lines, too.
                next_line = next_line.replace("#", "").strip()
                initialisation.description += "\n" + next_line
                lines[next_line_num] = ""

                line_num += 1

            workbook.initialisation.append(initialisation)

    def split_line(line):
        """
        Input of this function is a Test line, means it has VAR, INIT, and EV.
        This function splits the line into these 3 parts.
        :returns: Array with 3 elements: data[0]=Variable, data[1]=Initial-Value, data[2]=Expected-Value.
        """
        tmp_line = ""
        data = []
        for part in line.split(","):
            part = part.strip()
            if part.upper().startswith("INIT") or part.upper().startswith("EV"):
                data.append(tmp_line[:-1])
                tmp_line = ""
            tmp_line += part + ","
        data.append(tmp_line[:-1])

        # In case that INIT or EV is omitted
        if len(data) < 3:
            if data[1].lstrip().lower().startswith("init"):
                data.append("")
            elif data[1].lstrip().lower().startswith("ev"):
                ev = data[1]
                data[1] = ""
                data.append(ev)

        return data

    def extract_test_case_data(line_num):
        line = lines[line_num].strip()

        counter = line_num + 1
        while lines[counter].lstrip().startswith("--"):
            counter += 1

        if lines[counter].lstrip().startswith("&"):
            next_line = lines[counter].lstrip()
            while next_line.startswith("&"):
                line += next_line[1:].strip()
                counter += 1
                while lines[counter].lstrip().startswith("--"):
                    counter += 1
                next_line = lines[counter].lstrip()
        data = split_line(line)

        test_case = PtuWorkBook.TestCase()

        test_case.param_type = data[0].split()[0]
        test_case.param_name = data[0].replace(test_case.param_type, "").lstrip()

        if not test_case.param_name.split(",")[-1] == test_case.param_name:
            data_type = test_case.param_name.split(",")[-1]
            test_case.param_name = test_case.param_name.replace(data_type, "")
            test_case.param_name = test_case.param_name[:-1]
            if data_type.strip() != "":
                test_case.param_type += "," + data_type

        init = data[1]
        if init.count("=") == 1:
            init = init.replace("init", "")
            init = init.replace("INIT", "")
            init = init.replace("=", "")
            init = init.strip()
        test_case.init += init

        ev = data[2]
        if len(data) == 3:
            if ev.count("=") == 1:
                ev = ev.replace("ev", "")
                ev = ev.replace("EV", "")
                ev = ev.replace("=", "")
                ev = ev.strip()
        else:
            if ev.startswith("ev") or ev.startswith("EV"):
                delta = data[3].strip()
                delta = delta.replace("DELTA", "")
                delta = delta.replace("delta", "")
                delta = delta.replace("=", "±")
                delta = delta.strip()
                ev = ev.replace("ev", "")
                ev = ev.replace("EV", "")
                ev = ev.replace("=", "")
                ev = ev.strip()
                ev += " " + delta
            elif ev.startswith("MIN") or ev.startswith("min"):
                ev += data[3].rstrip()
        test_case.ev += ev

        test_case.conditions = check_for_conditions(line_num, line_num)

        return test_case

    def extract_environments_data():
        """
         Extracting Environments' test data by analyzing environments scopes
        """
        for counter in range(line_numbers.get_environment_count()):
            #Getting environment name
            environment_name = lines[line_numbers.ENVIRONMENT_START_LIST[counter]]
            environment_name = environment_name.replace("ENVIRONMENT ", "")

            #Creating new object of environment class for saving data
            environment = PtuWorkBook.Environment()
            environment.name = environment_name.strip()
            environment.conditions = check_for_conditions(line_numbers.ENVIRONMENT_START_LIST[counter],
                                                          line_numbers.ENVIRONMENT_END_LIST[counter])

            # Analyze inside the scope of the environment for test data
            for line_num in range(line_numbers.ENVIRONMENT_START_LIST[counter] + 1,
                                  line_numbers.ENVIRONMENT_END_LIST[counter]):
                line = lines[line_num].strip()

                if line.upper().startswith("VAR") or line.upper().startswith("ARRAY") or line.upper().startswith("STR"):
                    environment.test_case_list.append(extract_test_case_data(line_num))

            workbook.environments.append(environment)

    def extract_element_data(line_num, element, data_type_stack):
        line = lines[line_num].strip()
        if line.startswith("USE"):
            line = line.replace("USE", "")
            element.use.append(line.strip())
        elif line.startswith("--"):
            line = line.replace("--", "").strip().lower()
            if line.startswith("input"):
                data_type_stack.append("input")
            elif line.startswith("output"):
                data_type_stack.append("output")
            elif line.startswith("calib"):
                data_type_stack.append("calibrations")
        elif line.upper().startswith("VAR") or line.upper().startswith("ARRAY") or line.upper().startswith("STR"):
            test_case = extract_test_case_data(line_num)
            data_type = data_type_stack[-1]
            if data_type == "input":
                element.input_data.append(test_case)
            elif data_type == "output":
                element.output_data.append(test_case)
            elif data_type == "calibrations":
                element.calibrations.append(test_case)
            elif data_type == "":
                if "init" in test_case.ev or "INIT" in test_case.ev:
                    element.input_data.append(test_case)
                else:
                    element.output_data.append(test_case)
        elif line.startswith("STUB"):
            line = line.replace("STUB", "")
            stub = PtuWorkBook.Stub()
            stub.stub_definition += line.strip()
            stub.conditions = check_for_conditions(line_num, line_num)
            element.stub_calls.append(stub)
        elif line.startswith("#"):
            user_code = PtuWorkBook.UserCode()
            user_code.code += line[1:].strip()
            user_code.conditions = check_for_conditions(line_num, line_num)
            element.user_code.append(user_code)

    def extract_services_data():
        for service_line_num in line_numbers.SERVICE_LIST:
            service_name = lines[service_line_num.START]
            service_name = service_name.replace("SERVICE ", "")
            service = PtuWorkBook.Service()
            service.name = service_name.strip()
            service.conditions = check_for_conditions(service_line_num.START, service_line_num.END)

            for within_service_line_num in range(service_line_num.START, service_line_num.TEST_START[0]):
                line = lines[within_service_line_num]
                if line.startswith("#"):
                    user_code = PtuWorkBook.UserCode()
                    user_code.code += line[1:].strip()
                    user_code.conditions = check_for_conditions(within_service_line_num, within_service_line_num)
                    service.user_code.append(user_code)
                elif line.startswith("--"):
                    line = line.replace("--", "/* ") + " */"
                    user_code = PtuWorkBook.UserCode()
                    user_code.code += line.strip()
                    user_code.conditions = check_for_conditions(within_service_line_num, within_service_line_num)
                    service.user_code.append(user_code)
                elif line.startswith("USE"):
                    line = line.replace("USE", "")
                    service.use.append(line.strip())

            for counter in range(service_line_num.get_test_count()):
                test = PtuWorkBook.Test()
                for test_line_num in range(service_line_num.TEST_START[counter],
                                           service_line_num.ELEMENT_START[counter]):
                    line = lines[test_line_num].strip()
                    if line.upper().startswith("TEST"):
                        test.name = line
                    elif line.upper().startswith("FAMILY"):
                        family = line.replace("FAMILY", "")
                        test.family = family.strip()
                    elif line.upper().startswith("COMMENT"):
                        comment = line.replace("COMMENT", "")
                        test.comment.append(comment.strip())
                    elif line.upper().startswith("USE"):
                        use = line.replace("USE", "")
                        test.use.append(use.strip())
                element = PtuWorkBook.Element()
                data_type_stack = [""]
                for element_line_num in range(service_line_num.ELEMENT_START[counter],
                                              service_line_num.ELEMENT_END[counter]):
                    extract_element_data(element_line_num, element, data_type_stack)
                test.element = element
                test.conditions = check_for_conditions(service_line_num.TEST_START[counter],
                                                       service_line_num.TEST_END[counter])
                test.extend_use()
                service.test_list.append(test)
            service.extend_use()
            workbook.services.append(service)

    extract_preface_data()
    extract_include_data()
    extract_comment_data()
    extract_user_code_data()
    extract_stub_definitions_data()
    extract_initialization_data()
    extract_environments_data()
    extract_services_data()
