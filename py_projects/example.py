import csv
from docxtpl import DocxTemplate


def get_user_input():
    print("Введите данные для заполнения:")
    group_number = input("Номер группы: ")
    fio_student = input("ФИО студента: ")
    university_name = input("Название университета: ")
    faculty_number = input("Номер факультета: ")
    practice_place = input("Место практики: ")
    course_number = input("Номер курса: ")
    date_beginning = input("Дата начала (дд.мм.гггг): ")
    date_ending = input("Дата окончания (дд.мм.гггг): ")
    practice_director_fio = input("ФИО руководителя практики: ")
    practice_director_position = input("Должность руководителя практики: ").lower()
    departments_chair_name = input("ФИО заведующего кафедрой: ")
    practice_type_name = input("Тип практики (какой): ").lower()
    faculty_name = input("Название факультета: ")

    return {
        "group_number": group_number,
        "fio_student": fio_student,
        "university_name": university_name,
        "faculty_number": faculty_number,
        "practice_place": practice_place,
        "course_number": course_number,
        "date_beginning": date_beginning,
        "date_ending": date_ending,
        "practice_director_fio": practice_director_fio,
        "practice_director_position": practice_director_position,
        "departments_chair_name": departments_chair_name,
        "practice_type_name": practice_type_name,
        "faculty_name": faculty_name
    }


def save_to_csv(data, filename="data.csv"):
    with open(filename, mode='w', encoding='utf-8', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=data.keys())
        writer.writeheader()
        writer.writerow(data)


def generate_documents(csv_filename):
    with open(csv_filename, encoding='utf-8') as r_file:
        file_reader = csv.DictReader(r_file)

        for row in file_reader:
            doc1 = DocxTemplate("data/ИндЗадание.docx")
            doc2 = DocxTemplate("data/Заявление.docx")
            context = {
                "group_number": row["group_number"],
                "fio_student": row["fio_student"],
                "university_name": row["university_name"],
                "faculty_number": row["faculty_number"],
                "practice_place": row["practice_place"],
                "course_number": row["course_number"],
                "date_beginning": row["date_beginning"],
                "date_ending": row["date_ending"],
                "practice_director_fio": row["practice_director_fio"],
                "practice_director_position": row["practice_director_position"],
                "departments_chair_name": row["departments_chair_name"],
                "practice_type_name": row["practice_type_name"],
                "faculty_name": row["faculty_name"]
            }
            doc1.render(context)
            doc2.render(context)
            doc1.save(f"Индивидуальное задание {row['group_number']} {row['fio_student'].replace(' ', '_')}.docx")
            doc2.save(
                f"Заявление на прохождении практики {row['group_number']} {row['fio_student'].replace(' ', '_')}.docx")


def main():
    data = get_user_input()
    save_to_csv(data)
    generate_documents("data.csv")


if __name__ == "__main__":
    main()
