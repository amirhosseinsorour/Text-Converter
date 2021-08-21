from os import path, makedirs
import constants as const


class Line:
    """
    Objects of this class contain context of the line,
    number of line in old file and number of line in new file.
    """

    def __init__(self, context, old_line_num, new_line_num):
        self.context = context
        self.old_line_num = old_line_num
        self.new_line_num = new_line_num


class ErrorCatchStack:
    def __init__(self):
        """
        This class contains a stack for any type of error that may happen
        """
        self.if_stack = []
        self.define_stub_stack = []
        self.initialization_stack = []
        self.environment_stack = []
        self.service_stack = []
        self.test_stack = []
        self.element_stack = []

    """ Define some getters for length of error stacks """

    def get_if_stack_len(self):
        return len(self.if_stack)

    def get_define_stub_stack_len(self):
        return len(self.define_stub_stack)

    def get_initialization_stack_len(self):
        return len(self.initialization_stack)

    def get_environment_stack_len(self):
        return len(self.environment_stack)

    def get_service_stack_len(self):
        return len(self.service_stack)

    def get_test_stack_len(self):
        return len(self.test_stack)

    def get_element_stack_len(self):
        return len(self.element_stack)


def error_handler(file_handler, stack, error_type):
    """
    This function handles any error that may happen, due to error_type
    In simple words, this function adds missing endings for scopes with error.
    Implementation is not complete(handle by just raising exception)
    """
    if error_type == const.IF_ERROR:
        file_handler.seek(0, 0)
        all_lines = file_handler.readlines()
        file_handler.seek(0, 0)
        for i in range(len(all_lines) - 1):
            if all_lines[i + 1].startswith("##define const const"):
                for j in range(len(stack)):
                    file_handler.write("ENDIF\n")
            file_handler.write(all_lines[i])
    if error_type == const.DEFINE_STUB_ERROR:
        raise Exception("DEFINE STUB ERROR")
    if error_type == const.INITIALIZATION_ERROR:
        raise Exception("INITIALISATION ERROR")
    if error_type == const.ELEMENT_ERROR:
        raise Exception("ELEMENT ERROR")
    if error_type == const.SERVICE_ERROR:
        raise Exception("SERVICE ERROR")
    if error_type == const.TEST_ERROR:
        raise Exception("TEST ERROR")
    if error_type == const.ELEMENT_ERROR:
        raise Exception("ELEMENT ERROR")


def handle_all_errors(error_stack, new_file):
    """
    This function checks if an error has occurred by checking length of error stacks
    :param error_stack: obj of ErrorCatchStack class,
    :param new_file: file handler of new file in which data must be written
    :return: None
    """

    # Check for IF ERROR
    if error_stack.get_if_stack_len() > 0:
        error_handler(new_file, error_stack.if_stack, const.IF_ERROR)

    # Check for DEFINE STUB ERROR
    if error_stack.get_define_stub_stack_len() > 0:
        error_handler(new_file, error_stack.define_stub_stack, const.DEFINE_STUB_ERROR)

    # Check for INITIALIZATION ERROR
    if error_stack.get_initialization_stack_len() > 0:
        error_handler(new_file, error_stack.initialization_stack, const.INITIALIZATION_ERROR)

    # Check for ENVIRONMENT ERROR
    if error_stack.get_environment_stack_len() > 0:
        error_handler(new_file, error_stack.environment_stack, const.ENVIRONMENT_ERROR)

    # Check for SERVICE ERROR
    if error_stack.get_service_stack_len() > 0:
        error_handler(new_file, error_stack.service_stack, const.SERVICE_ERROR)

    # Check for TEST ERROR
    if error_stack.get_test_stack_len() > 0:
        error_handler(new_file, error_stack.test_stack, const.TEST_ERROR)

    # Check for ELEMENT ERROR
    if error_stack.get_element_stack_len() > 0:
        error_handler(new_file, error_stack.element_stack, const.ELEMENT_ERROR)


def check_for_errors(error_stack, line):
    """
    This function checks for any type of error that may happen.
    Errors happen when scope of sth has been started and there is no ending for it.
    In these cases, an ending must be added for that scope.
    :param line: is checked to be starter of or ending for a scope.
    If it was a starter of a scope, it is pushed to a stacked until finding it's ending.
    If it was an ending for a scope, starter of that scope is popped out from the stack.
    :param error_stack: obj of ErrorCatchStack class, containing a stack for any type of error.
    """

    line_context = line.context.strip().upper()

    # Check if the line has started scope of an IF. If so, push it into stack
    if line_context.startswith("IF"):
        error_stack.if_stack.append(line)
        # print([x.context + str(x.old_line_num) for x in error_stack.if_stack])
    # Check if the line has ended scope of an IF. If so, pop the last element from stack
    elif line_context.startswith("ENDIF") or line_context.startswith("END IF"):
        error_stack.if_stack.pop()
        # print([x.context + str(x.old_line_num) for x in error_stack.if_stack])

    # Check if the line has started scope of an ELEMENT. If so, push it into stack
    elif line_context.startswith("ELEMENT"):
        error_stack.element_stack.append(line)
    # Check if the line has ended scope of an ELEMENT. If so, pop the last element from stack
    elif line_context.startswith("END ELEMENT"):
        error_stack.element_stack.pop()

    # Check if the line has started scope of a TEST. If so, push it into stack
    elif line_context.startswith("TEST"):
        error_stack.test_stack.append(line)
    # Check if the line has ended scope of a TEST. If so, pop the last element from stack
    elif line_context.startswith("END TEST"):
        error_stack.test_stack.pop()

    # Check if the line has started scope of a SERVICE. If so, push it into stack
    elif line_context.startswith("SERVICE "):
        error_stack.service_stack.append(line)
    # Check if the line has ended scope of a SERVICE. If so, pop the last element from stack
    elif line_context.startswith("END SERVICE"):
        error_stack.service_stack.pop()

    # Check if the line has started scope of an ENVIRONMENT. If so, push it into stack
    elif line_context.startswith("ENVIRONMENT"):
        error_stack.environment_stack.append(line)
    # Check if the line has ended scope of an ENVIRONMENT. If so, pop the last element from stack
    elif line_context.startswith("END ENVIRONMENT"):
        error_stack.environment_stack.pop()

    # Check if the line has started scope of a STUB DEFINITION. If so, push it into stack
    elif line_context.startswith("DEFINE STUB"):
        error_stack.define_stub_stack.append(line)
    # Check if the line has ended scope of a STUB DEFINITION. If so, pop the last element from stack
    elif line_context.startswith("END DEFINE"):
        error_stack.define_stub_stack.pop()

    # Check if the line has started scope of INITIALISATION. If so, push it into stack
    elif line_context.startswith("INITIALISATION"):
        error_stack.initialization_stack.append(line)
    # Check if the line has ended scope of INITIALISATION. If so, pop the last element from stack
    elif line_context.startswith("END INITIALISATION"):
        error_stack.initialization_stack.pop()


def writable_comment(comment_line):
    """
    This function is used to check if a comment should be written into new file or not.
     :returns False if not

     If comment includes just additional characters, it is useless to write it.
     Example : " --~T "
    """
    if "--~" in comment_line and len(comment_line) == 5:
        return False
    return True


def rewrite_ptu_file(file_name, old_path, new_path):
    """
    This function opens ptu file and rewrite it into new file.
    New file is now easier and more simple to read.
    Also, if there is a scope in ptu file with no ending, it is handled in new file.
    :param file_name: name of ptu file
    :param old_path: path of ptu file to be read
    :param new_path: path of new file to be written into
    :return: None
    """

    """ Open old file for reading """

    old_file = open(old_path + "\\" + file_name, "rt")
    lines = old_file.readlines()  # Reading all lines and save in an array
    old_file.close()

    """ Open new file for writing """

    if not path.exists(new_path):
        makedirs(new_path)  # If folder doesn't exist, create it

    new_file = open(new_path + "\\" + file_name, "w+t")
    new_line_num = 0  # Counter for keeping number of line which is being written in new file

    error_stack = ErrorCatchStack()  # Stack for detecting start and end of scopes and handle if an error occurs

    # Reading old file lines and write it in new file, just in case
    for old_line_num in range(len(lines)):

        # Create object of Line class for every line that is being read
        line = Line(lines[old_line_num], old_line_num, new_line_num)

        # If the line starts or ends scope of sth, it must be checked due to handling errors that may happen
        check_for_errors(error_stack, line)

        if line.context.startswith("--"):  # If line is comment
            if writable_comment(line.context):  # If the comment is ok to be written in the new file
                if "--~+:" in line.context:  # Check comment for additional characters
                    line.context = line.context.replace("~+:", " ")  # Omitting additional characters
            else:
                # If comment should not be written into new file, continue loop and do not increase
                # new_line_num counter. So nothing will be written into new file
                continue

        # Writing file name in new file and modify it in case it is not correct or is corrupted
        if line.context.startswith("-- Filename"):
            line.context = line.context.replace(".ptu", "")  # Omitting extension
            line.context = line.context.replace("template", file_name.replace(".ptu", ""))

        # After checking everything, write the line into new file and increase new_line_num counter
        new_file.write(line.context)
        new_line_num += 1

    # Now that everything is checked, detect if an error has occurred and handle it
    handle_all_errors(error_stack, new_file)

    new_file.close()
