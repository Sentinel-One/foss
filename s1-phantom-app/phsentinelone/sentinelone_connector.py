# -----------------------------------------
# Phantom sample App Connector python file
# Lee Wei (leewei@sentinelone.com)
# -----------------------------------------

# Phantom App imports
import phantom.app as phantom
from phantom.base_connector import BaseConnector
from phantom.action_result import ActionResult

# Usage of the consts file is recommended
# from sentinelone_consts import *
import requests
import json
from bs4 import BeautifulSoup


class RetVal(tuple):
    def __new__(cls, val1, val2):
        return tuple.__new__(RetVal, (val1, val2))


class SentineloneConnector(BaseConnector):

    def __init__(self):

        # Call the BaseConnectors init first
        super(SentineloneConnector, self).__init__()

        self._state = None

        # Variable to hold a base_url in case the app makes REST calls
        # Do note that the app json defines the asset config, so please
        # modify this as you deem fit.
        self._base_url = None
        self.HEADER = {"Content-Type": "application/json"}

    def _process_empty_reponse(self, response, action_result):

        if response.status_code == 200:
            return RetVal(phantom.APP_SUCCESS, {})

        return RetVal(action_result.set_status(phantom.APP_ERROR, "Empty response and no information in the header"), None)

    def _process_html_response(self, response, action_result):

        # An html response, treat it like an error
        status_code = response.status_code

        try:
            soup = BeautifulSoup(response.text, "html.parser")
            error_text = soup.text
            split_lines = error_text.split('\n')
            split_lines = [x.strip() for x in split_lines if x.strip()]
            error_text = '\n'.join(split_lines)
        except:
            error_text = "Cannot parse error details"

        message = "Status Code: {0}. Data from server:\n{1}\n".format(status_code,
                error_text)

        message = message.replace('{', '{{').replace('}', '}}')

        return RetVal(action_result.set_status(phantom.APP_ERROR, message), None)

    def _process_json_response(self, r, action_result):

        # Try a json parse
        try:
            resp_json = r.json()
        except Exception as e:
            return RetVal(action_result.set_status(phantom.APP_ERROR, "Unable to parse JSON response. Error: {0}".format(str(e))), None)

        # Please specify the status codes here
        if 200 <= r.status_code < 399:
            return RetVal(phantom.APP_SUCCESS, resp_json)

        # You should process the error returned in the json
        message = "Error from server. Status Code: {0} Data from server: {1}".format(
                r.status_code, r.text.replace('{', '{{').replace('}', '}}'))

        return RetVal(action_result.set_status(phantom.APP_ERROR, message), None)

    def _process_response(self, r, action_result):

        # store the r_text in debug data, it will get dumped in the logs if the action fails
        if hasattr(action_result, 'add_debug_data'):
            action_result.add_debug_data({'r_status_code': r.status_code})
            action_result.add_debug_data({'r_text': r.text})
            action_result.add_debug_data({'r_headers': r.headers})

        # Process each 'Content-Type' of response separately

        # Process a json response
        if 'json' in r.headers.get('Content-Type', ''):
            return self._process_json_response(r, action_result)

        # Process an HTML resonse, Do this no matter what the api talks.
        # There is a high chance of a PROXY in between phantom and the rest of
        # world, in case of errors, PROXY's return HTML, this function parses
        # the error and adds it to the action_result.
        if 'html' in r.headers.get('Content-Type', ''):
            return self._process_html_response(r, action_result)

        # it's not content-type that is to be parsed, handle an empty response
        if not r.text:
            return self._process_empty_reponse(r, action_result)

        # everything else is actually an error at this point
        message = "Can't process response from server. Status Code: {0} Data from server: {1}".format(
                r.status_code, r.text.replace('{', '{{').replace('}', '}}'))

        return RetVal(action_result.set_status(phantom.APP_ERROR, message), None)

    def _make_rest_call(self, endpoint, action_result, headers=None, params=None, data=None, method="get"):

        config = self.get_config()

        resp_json = None

        try:
            request_func = getattr(requests, method)
        except AttributeError:
            return RetVal(action_result.set_status(phantom.APP_ERROR, "Invalid method: {0}".format(method)), resp_json)

        # Create a URL to connect to
        url = self._base_url + endpoint

        self.save_progress(url)
        try:
            r = request_func(
                            url,
                            # auth='Token ' + config.get('access_token', False),
                            json=data,
                            headers=headers,
                            verify=config.get('verify_server_cert', False),
                            params=params)
        except Exception as e:
            return RetVal(action_result.set_status( phantom.APP_ERROR, "Error Connecting to server. Details: {0}".format(str(e))), resp_json)

        return self._process_response(r, action_result)

    def _handle_test_connectivity(self, param):

        # Add an action result object to self (BaseConnector) to represent the action for this param
        action_result = self.add_action_result(ActionResult(dict(param)))

        self.save_progress("Connecting to the SentinelOne server")

        # make rest call
        header = self.HEADER
        header["Authorization"] = "Token %s" % self.token
        ret_val, response = self._make_rest_call('/web/api/v1.6/threats/summary', action_result, headers=header)
        self.save_progress("response: {0}".format(response))

        if (phantom.is_fail(ret_val)):
            # the call to the 3rd party device or service failed, action result should contain all the error details
            # so just return from here
            self.save_progress("Test Connectivity Failed. Error: {0}".format(action_result.get_message()))
            return action_result.get_status()

        # Return success
        self.save_progress("Login to SentinelOne server is successful")
        return action_result.set_status(phantom.APP_SUCCESS)

    def _handle_block_hash(self, param):

        self.save_progress("In action handler for: {0}".format(self.get_action_identifier()))
        action_result = self.add_action_result(ActionResult(dict(param)))
        hash = param['hash']
        description = param['description']
        os_family = param['os_family']

        summary = action_result.update_summary({})
        summary['hash'] = hash
        summary['description'] = description

        header = self.HEADER
        header["Authorization"] = "Token %s" % self.token
        header["Content-Type"] = "application/json"
        body = {"description": description, "os_family": os_family, "is_black": True, "hash": hash}

        ret_val, response = self._make_rest_call('/web/api/v1.6/hashes', action_result, headers=header, method='post', data=body)
        self.save_progress("response: {0}".format(response))

        if (phantom.is_fail(ret_val)):
            return action_result.get_status()

        return action_result.set_status(phantom.APP_SUCCESS)

    def _handle_unblock_hash(self, param):

        self.save_progress("In action handler for: {0}".format(self.get_action_identifier()))
        action_result = self.add_action_result(ActionResult(dict(param)))
        hash = param['hash']

        summary = action_result.update_summary({})
        summary['hash'] = hash

        header = self.HEADER
        header["Authorization"] = "Token %s" % self.token
        header["Content-Type"] = "application/json"
        ret_val, response = self._make_rest_call('/web/api/v1.6/hashes/' + hash, action_result, headers=header, method='delete')
        self.save_progress("response: {0}".format(response))

        # if (phantom.is_fail(ret_val)):
        #    return action_result.get_status()

        return action_result.set_status(phantom.APP_SUCCESS)

    def _handle_list_endpoints(self, param):

        # Implement the handler here
        # use self.save_progress(...) to send progress messages back to the platform
        self.save_progress("In action handler for: {0}".format(self.get_action_identifier()))

        # Add an action result object to self (BaseConnector) to represent the action for this param
        action_result = self.add_action_result(ActionResult(dict(param)))

        # make rest call
        header = self.HEADER
        header["Authorization"] = "Token %s" % self.token
        ret_val, response = self._make_rest_call('/web/api/v1.6/agents', action_result, headers=header)
        self.save_progress("ret_val: {0}".format(ret_val))
        self.save_progress("response: {0}".format(response))

        if (phantom.is_fail(ret_val)):
            # the call to the 3rd party device or service failed, action result should contain all the error details
            # so just return from here
            return action_result.get_status()

        # Now post process the data,  uncomment code as you deem fit

        # Add the response into the data section
        action_result.add_data(response)

        # Add a dictionary that is made up of the most important values from data into the summary
        # summary = action_result.update_summary({})
        # summary['computer_name'] = response["network_information"]["computer_name"]

        # Return success, no need to set the message, only the status
        # BaseConnector will create a textual message based off of the summary dictionary
        return action_result.set_status(phantom.APP_SUCCESS)

    def _handle_list_processes(self, param):

        self.save_progress("In action handler for: {0}".format(self.get_action_identifier()))
        action_result = self.add_action_result(ActionResult(dict(param)))

        ip_hostname = param['ip_hostname']
        ret_val = self._get_agent_id(ip_hostname, action_result)
        self.save_progress('Agent query: ' + ret_val)

        if (ret_val == '0'):
            return action_result.set_status(phantom.APP_SUCCESS, "Endpoint not found")
        elif (ret_val == '99'):
            return action_result.set_status(phantom.APP_SUCCESS, "More than one endpoint found")
        else:
            summary = action_result.update_summary({})
            summary['ip_hostname'] = ip_hostname
            summary['agent_id'] = ret_val

            # make rest call
            # GET /web/api/v1.6/agents/{id}/processes
            header = self.HEADER
            header["Authorization"] = "Token %s" % self.token
            ret_val, response = self._make_rest_call('/web/api/v1.6/agents/' + ret_val + '/processes', action_result, headers=header)
            self.save_progress("ret_val: {0}".format(ret_val))
            self.save_progress("response: {0}".format(response))

        # if (phantom.is_fail(ret_val)):
        #    return action_result.get_status()

        action_result.add_data(response)
        return action_result.set_status(phantom.APP_SUCCESS)

    def _handle_quarantine_device(self, param):

        self.save_progress("In action handler for: {0}".format(self.get_action_identifier()))
        action_result = self.add_action_result(ActionResult(dict(param)))
        ip_hostname = param['ip_hostname']

        ret_val = self._get_agent_id(ip_hostname, action_result)

        self.save_progress('Agent query: ' + ret_val)

        if (ret_val == '0'):
            return action_result.set_status(phantom.APP_SUCCESS, "Endpoint not found")
        elif (ret_val == '99'):
            return action_result.set_status(phantom.APP_SUCCESS, "More than one endpoint found")
        else:
            summary = action_result.update_summary({})
            summary['ip_hostname'] = ip_hostname
            summary['agent_id'] = ret_val

            # /web/api/v1.6/agents/{id}/disconnect
            header = self.HEADER
            header["Authorization"] = "Token %s" % self.token
            ret_val, response = self._make_rest_call('/web/api/v1.6/agents/' + ret_val + '/disconnect', action_result, headers=header, method='post')
            self.save_progress("response: {0}".format(response))

        return action_result.set_status(phantom.APP_SUCCESS)

    def _handle_unquarantine_device(self, param):

        self.save_progress("In action handler for: {0}".format(self.get_action_identifier()))
        action_result = self.add_action_result(ActionResult(dict(param)))
        ip_hostname = param['ip_hostname']

        ret_val = self._get_agent_id(ip_hostname, action_result)

        self.save_progress('Agent query: ' + ret_val)

        if (ret_val == '0'):
            return action_result.set_status(phantom.APP_SUCCESS, "Endpoint not found")
        elif (ret_val == '99'):
            return action_result.set_status(phantom.APP_SUCCESS, "More than one endpoint found")
        else:
            summary = action_result.update_summary({})
            summary['ip_hostname'] = ip_hostname
            summary['agent_id'] = ret_val

            # /web/api/v1.6/agents/{id}/connect
            header = self.HEADER
            header["Authorization"] = "Token %s" % self.token
            ret_val, response = self._make_rest_call('/web/api/v1.6/agents/' + ret_val + '/connect', action_result, headers=header, method='post')
            self.save_progress("response: {0}".format(response))

        return action_result.set_status(phantom.APP_SUCCESS)

    def _handle_scan_endpoint(self, param):

        self.save_progress("In action handler for: {0}".format(self.get_action_identifier()))
        action_result = self.add_action_result(ActionResult(dict(param)))
        ip_hostname = param['ip_hostname']

        ret_val = self._get_agent_id(ip_hostname, action_result)

        self.save_progress('Agent query: ' + ret_val)

        if (ret_val == '0'):
            return action_result.set_status(phantom.APP_SUCCESS, "Endpoint not found")
        elif (ret_val == '99'):
            return action_result.set_status(phantom.APP_SUCCESS, "More than one endpoint found")
        else:
            summary = action_result.update_summary({})
            summary['ip_hostname'] = ip_hostname
            summary['agent_id'] = ret_val

            header = self.HEADER
            header["Authorization"] = "Token %s" % self.token
            ret_val, response = self._make_rest_call('/web/api/v1.6/agents/' + ret_val + '/initiate-scan', action_result, headers=header, method='post')
            self.save_progress("response: {0}".format(response))

        return action_result.set_status(phantom.APP_SUCCESS)

    def _handle_get_endpoint_info(self, param):

        self.save_progress("In action handler for: {0}".format(self.get_action_identifier()))
        action_result = self.add_action_result(ActionResult(dict(param)))

        ip_hostname = param['ip_hostname']
        ret_val = self._get_agent_id(ip_hostname, action_result)
        self.save_progress('Agent query: ' + ret_val)

        if (ret_val == '0'):
            return action_result.set_status(phantom.APP_SUCCESS, "Endpoint not found")
        elif (ret_val == '99'):
            return action_result.set_status(phantom.APP_SUCCESS, "More than one endpoint found")
        else:
            summary = action_result.update_summary({})
            summary['ip_hostname'] = ip_hostname
            summary['agent_id'] = ret_val

            # make rest call
            # GET /web/api/v1.6/agents/{id}
            header = self.HEADER
            header["Authorization"] = "Token %s" % self.token
            ret_val, response = self._make_rest_call('/web/api/v1.6/agents/' + ret_val, action_result, headers=header)
            self.save_progress("ret_val: {0}".format(ret_val))
            self.save_progress("response: {0}".format(response))

        # if (phantom.is_fail(ret_val)):
        #    return action_result.get_status()

        action_result.add_data(response)
        return action_result.set_status(phantom.APP_SUCCESS)

    def _handle_mitigate_threat(self, param):

        self.save_progress("In action handler for: {0}".format(self.get_action_identifier()))
        action_result = self.add_action_result(ActionResult(dict(param)))
        threat_id = param['threat_id']
        action = param['action']

        summary = action_result.update_summary({})
        summary['threat_id'] = threat_id
        summary['action'] = action

        header = self.HEADER
        header["Authorization"] = "Token %s" % self.token
        header["Content-Type"] = "application/json"

        # POST /web/api/v1.6/threats/:threat_id/mitigate/:action
        ret_val, response = self._make_rest_call('/web/api/v1.6/threats/' + threat_id + '/mitigate/' + action, action_result, headers=header, method='post')
        self.save_progress("response: {0}".format(response))

        # if (phantom.is_fail(ret_val)):
        #    return action_result.get_status()

        return action_result.set_status(phantom.APP_SUCCESS)

    def is_empty(self, any_structure):
        if any_structure:
            return False
        else:
            return True

    def _get_agent_id(self, search_text, action_result):
        # First lookup the Agent ID
        header = self.HEADER
        header["Authorization"] = "Token %s" % self.token
        ret_val, response = self._make_rest_call('/web/api/v1.6/agents?query=' + search_text, action_result, headers=header)
        self.save_progress("response: {0}".format(response))

        if (phantom.is_fail(ret_val)):
            return str(-1)

        endpoints_found = len(response)
        self.save_progress("Endpoints found: " + str(endpoints_found))
        action_result.add_data(response)

        if (endpoints_found == 0):
            return '0'
        elif (endpoints_found > 1):
            return '99'
        else:
            return response[0]['id']

    def handle_action(self, param):

        ret_val = phantom.APP_SUCCESS

        # Get the action that we are supposed to execute for this App Run
        action_id = self.get_action_identifier()

        self.debug_print("action_id", self.get_action_identifier())

        if action_id == 'test_connectivity':
            ret_val = self._handle_test_connectivity(param)

        elif action_id == 'block_hash':
            ret_val = self._handle_block_hash(param)

        elif action_id == 'list_endpoints':
            ret_val = self._handle_list_endpoints(param)

        elif action_id == 'list_processes':
            ret_val = self._handle_list_processes(param)

        elif action_id == 'quarantine_device':
            ret_val = self._handle_quarantine_device(param)

        elif action_id == 'unquarantine_device':
            ret_val = self._handle_unquarantine_device(param)

        elif action_id == 'unblock_hash':
            ret_val = self._handle_unblock_hash(param)

        elif action_id == 'mitigate_threat':
            ret_val = self._handle_mitigate_threat(param)

        elif action_id == 'scan_endpoint':
            ret_val = self._handle_scan_endpoint(param)

        elif action_id == 'get_endpoint_info':
            ret_val = self._handle_get_endpoint_info(param)

        return ret_val

    def initialize(self):

        # Load the state in initialize, use it to store data
        # that needs to be accessed across actions
        self._state = self.load_state()

        # get the asset config
        config = self.get_config()

        # Access values in asset config by the name

        # Required values can be accessed directly
        self._base_url = config['sentinelone_server_url']

        # Optional values should use the .get() function
        self.token = config.get('access_token')
        # optional_config_name = config.get('optional_config_name')

        return phantom.APP_SUCCESS

    def finalize(self):

        # Save the state, this data is saved accross actions and app upgrades
        self.save_state(self._state)
        return phantom.APP_SUCCESS


if __name__ == '__main__':

    import sys
    import pudb
    import argparse

    pudb.set_trace()

    argparser = argparse.ArgumentParser()

    argparser.add_argument('input_test_json', help='Input Test JSON file')
    argparser.add_argument('-u', '--username', help='username', required=False)
    argparser.add_argument('-p', '--password', help='password', required=False)

    args = argparser.parse_args()
    session_id = None

    username = args.username
    password = args.password

    if (username is not None and password is None):

        # User specified a username but not a password, so ask
        import getpass
        password = getpass.getpass("Password: ")

    if (username and password):
        try:
            print ("Accessing the Login page")
            r = requests.get("https://127.0.0.1/login", verify=False)
            csrftoken = r.cookies['csrftoken']

            data = dict()
            data['username'] = username
            data['password'] = password
            data['csrfmiddlewaretoken'] = csrftoken

            headers = dict()
            headers['Cookie'] = 'csrftoken=' + csrftoken
            headers['Referer'] = 'https://127.0.0.1/login'

            print ("Logging into Platform to get the session id")
            r2 = requests.post("https://127.0.0.1/login", verify=False, data=data, headers=headers)
            session_id = r2.cookies['sessionid']
        except Exception as e:
            print ("Unable to get session id from the platfrom. Error: " + str(e))
            exit(1)

    if (len(sys.argv) < 2):
        print "No test json specified as input"
        exit(0)

    with open(sys.argv[1]) as f:
        in_json = f.read()
        in_json = json.loads(in_json)
        print(json.dumps(in_json, indent=4))

        connector = SentineloneConnector()
        connector.print_progress_message = True

        if (session_id is not None):
            in_json['user_session_token'] = session_id

        ret_val = connector._handle_action(json.dumps(in_json), None)
        print (json.dumps(json.loads(ret_val), indent=4))

    exit(0)
