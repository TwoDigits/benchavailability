from flask import Flask, request, jsonify
import bench_availability_reminder as reminder

app = Flask(__name__)

# Route for triggering the reminder for bench candidates
# The bench availibilty list is provided as an Excel file
# The reminder will be sent as an Email to the candidates who are currently on the bench


@app.route('/api/v1/remind_candidates', methods=['POST'])
def remind_bench_candidates():
    availability_list_source_folder_url = request.form['availability_list_source_folder_url']
    if availability_list_source_folder_url is None:
        return jsonify({'Error': 'Availability list sourec folder not provided'}), 400
    reminder.check_availabilities_and_send_reminder(availability_list_source_folder_url)

    resp = jsonify(success=True)
    resp.status_code = 201
    return resp


if __name__ == '__main__':
    app.run()
