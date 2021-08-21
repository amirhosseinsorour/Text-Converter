from string import digits
from os import path, makedirs
import extract_data


def convert(ptu_obj: extract_data.PtuWorkBook, output_path) -> None:
    def write_data():
        def write_line(string=""):
            tst_file.write(string + "\n")

        def write_loop(i, length):
            length = length[length.find("TAB"): length.find("-")]
            loop_str = "for (" + i + "=0 ; " + i + "<" + length + " ; " + i + "++){"
            write_line(loop_str)

        def write_input_data(param, value):
            write_line("\t" + param + " = " + value + " ;")

        def write_output_data(param, value):
            write_line("\t{{ " + param + " == " + value + " }}")

        def write_service_call(param, service_name, inputs):
            write_line("\t" + param + " = " + service_name + "(" + ",".join(inputs) + ") ;")

        def write_test_data(element):
            input_data = element.input_data
            output_data = element.output_data
            input_params = []
            in_loop = False
            loop_var = ""

            for data in input_data:
                if data.init.upper().startswith("INIT FROM"):
                    write_loop(data.param_name, data.init)
                    in_loop = True
                    loop_var = data.param_name
                else:
                    if in_loop and "[" + loop_var + "]" not in data.init:
                        write_line("}")
                        in_loop = False
                    write_input_data(data.param_name, data.init)
                    input_params.append(data.param_name)

            for data in output_data:
                if in_loop and "[" + loop_var + "]" not in data.ev:
                    write_line("}")
                    in_loop = False
                write_service_call(data.param_name, service.name, input_params)
                write_output_data(data.param_name, data.ev)

            if in_loop:
                write_line("}")

        def write_user_code_data(element):
            write_line("TEST.VALUE_USER_CODE:<<testcase>>")
            write_line(service.get_all_user_code())
            write_test_data(element)
            write_line("TEST.END_VALUE_USER_CODE:")

        def write_test_case_data(test: extract_data.PtuWorkBook.Test):
            write_line("TEST.NEW")
            write_line("TEST.NAME: " + test.name)
            write_user_code_data(test.element)
            write_line("TEST.END")
            write_line()

        def write_subprogram_data(service: extract_data.PtuWorkBook.Service):
            write_line()
            unit_name = file_name.translate(file_name.maketrans('', '', digits)) + "_Ccode"
            write_line("TEST.UNIT:" + unit_name)
            write_line("TEST.SUBPROGRAM:" + service.name)
            write_line()
            for test in service.test_list:
                write_test_case_data(test)

        for service in ptu_obj.services:
            write_subprogram_data(service)

    file_name = ptu_obj.name
    if not path.exists(output_path):
        makedirs(output_path)  # If folder doesn't exist, create it
    tst_file = open(output_path + "\\" + file_name + ".tst", "wt")
    write_data()
    tst_file.close()
