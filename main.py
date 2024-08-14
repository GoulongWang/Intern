import splunklib.client as client
import splunklib.results as results
def connect_to_splunk(username,password,host='10.63.131.62',port='8089',owner='admin',app='search',sharing='user'):
    try:
        service = client.connect(host=host, port=port,username=username, password=password,owner=owner,app=app,sharing=sharing)
        if service:
            print("Splunk service created successfully")
            print("------------------------------------")
        return service
    except Exception as e:
        print(e)
def run_normal_mode_search(splunk_service,search_string,payload={}):
    try:
        job = splunk_service.jobs.create(search_string,**payload)
        while True:
            while not job.is_ready():
                pass
            if job["isDone"] == "1":

                break
        #print(job.content)
        print(job.results())
        '''
        
        reader = results.JSONResultsReader(job.results())
        for item in reader:
            print("分隔")
            print(item)
        '''
        '''
        for result in results.ResultsReader(job.results()):
            print(result)

        '''

    except Exception as e:
        print(e)

def main():
    try:
        splunk_service = connect_to_splunk(username='w9015431',password='!123qWeAsZ')
        search_string= ("search index=_audit (info=granted OR info=completed) "
                        "savedsearch_name=\"Checkpoint內對外阻擋TOP50: daily report\" "
                        "search_id=\"*scheduler*\" "
                        "| head 1"
                        "| table search_id, savedsearch_name, _time")
        payload = {"exec_mode": "normal", "earliest_time": "-24h", "latest_time": "now"}
        run_normal_mode_search(splunk_service, search_string, payload)
        print("成功拉")
    except Exception as e:
        print(e)

if __name__ == "__main__":
    main()