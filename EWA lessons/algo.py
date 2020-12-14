import datetime
import get_spreadsheet_data
from matplotlib import pyplot as plt

KEYS_LIST = {
    'Любов Федорко': '1C0I0fdWnM_qsyIh0k3_Sdm34VuZcqiBqz7wMS4a6dS0',
    'Альона Менчинська': '1KJIL16lIwLKXZC6gPa_I8PwniuPHwygnRK-Uhg2BB9s',
    'Вікторія Мезіна': '1-J2NE2wHGqWfoA5Tb9sb-fEm97QZXE8M3F7qQi3BkIU',
    'Оксана Кулик': '1YLRJM8zvk6CNnsUPMYw_M-0t6PMYoPMNY1-nAYxuKgU',
    'Вікторія Складанюк': '1BBgO05UGC_DE9yX2j5YKuormQIDZhocqzpxSlMvhx2Q',
    'Софія Аль-Лахам': '1ZYFsN4Ci515YUm7zDA0pjHJnRU9FYDqafldzanwuUZc',
    'Наталія Гусар': '1win9NmY0DMHUiFIppnyAaE4LSk8VKXHBa4w_ovDUHE0',
    'Лілія Сьомер': '1HxVAkwl4GpQaO5W_XfEgelm_foWrzL22a0MoO9sXJoQ',
    'Наталія Юрків ': '1m5p_DibJpmff4fTJ5DeylZwgcYfPYTmsmUjUjCNkhnY',
    'Ірина Гресько': '1-MHRYkq-iQqtOm5ilJIPCUbmQUVi7aVRGX7fZQMrq68',
    'Євгенія Іващенко': '1GLzedmQJc7bjf9gyPUSA1FrT55vrxxyYoYCYmeBKlYo',
    'Софія Дудлій': '1_Sx3vzWrWxAy1ip9d_8eFuY9KND1PGLg9ggSTkUSA9c',
    'Марія Петращук ': '1S6GzM_nQOYuUee8folQ3tcbbIBc4XipVumujxm_zQXk',
    'Оксана Кулік': '1RfjNad--S5T_DgwGr1yXDBLnIfJ1tdKcrICZ6XG_WDM',
    'Лілія Продоус': '1RkIZQ9Qhy6kuuiwcMZxfoPNDDvXqbZ9qxJjaOn2W6gY',
    'Катерина Назаренко': '1nqqgT0WAbORvq6oHC2p4-afN87v8ZL-gDxQQmDkNr_o',
}


class Excel:
    def __init__(self, values, teacher_name):
        self.values = values
        self.lessons_count = 0
        self.teacher_name = teacher_name
        self.teacher_lessons = {}
        self.all_dates_array = []
        if teacher_name == 'Лілія Продоус':
            self.sheet_start = 2
        else:
            self.sheet_start = 4

    def bonuses(self):

        pass
        #get data
        #count zp for all month
        #count retention
        #count intens
        #add persents
        #

    def retention(self):
        lessons_count = 7
        days_count = 30
        firs_lesson_day = 3
        retention = lessons_count / ((days_count - firs_lesson_day) * 8 / days_count)
        print(retention)


    @staticmethod
    def convert_xls_datetime(xls_date):
        return (datetime.datetime(1899, 12, 30)
                + datetime.timedelta(days=xls_date)).date()

    @staticmethod
    def lesson_cost_nazarenko(student_level):
        level = student_level.replace(' ', '').lower()
        lcs = 100
        if level == 'elementary' or level == 'pre-intermediate':
            lcs = 115

        elif level == 'intermediate':
            lcs = 140

        elif level == 'upper-inermediate' or level == 'upper-intermediate' or level == 'advanced':
            lcs = 165

        return lcs

    @staticmethod
    def lesson_cost_prodous(student_level, column_count):
        level = student_level.replace(' ', '').lower()
        lcs = 100
        if level == 'elementary' or level == 'pre-intermediate':
            if column_count > 27:
                lcs = 110
            else:
                lcs = 100

        elif level == 'intermediate':
            if column_count > 27:
                lcs = 130
            else:
                lcs = 120

        elif level == 'upper-inermediate' or level == 'upper-intermediate' or level == 'advanced':
            if column_count > 27:
                lcs = 150
            else:
                lcs = 140

        return lcs

    def get_lesson_sum(self,):
        today_date = datetime.date.today()
        if 1 <= today_date.day <= 15:
            date1 = datetime.date(today_date.year, today_date.month, 16)
            date2 = datetime.date(today_date.year, today_date.month, 30)
        print(today_date)
        # Main loop where we getting data from each line at sheet
        lesson_cost = 100
        for key in self.values:
            tmp_list = []

            # taking dates from line
            for column_count in range(self.sheet_start, len(key)):

                # looking for date
                if type(key[column_count]) is str and len(key[column_count]) > 5:
                    if self.teacher_name == 'Лілія Продоус':
                        lesson_cost = self.lesson_cost_prodous(key[column_count], column_count)
                    elif self.teacher_name == 'Катерина Назаренко':
                        lesson_cost = self.lesson_cost_nazarenko(key[column_count])

                elif type(key[column_count]) is int and key[column_count] > 10000:
                    date = self.convert_xls_datetime(key[column_count])
                    self.all_dates_array.append(date)

                    # if date at range - add to count
                    if date1 <= date <= date2:
                        tmp_list.append(date)
                        self.lessons_count += lesson_cost

            # print("Уроки: ", tmp_list, "\n")
        print('Кількість уроків за період ', date1, '--', date2, 'Sum:', self.lessons_count)
        print(self.teacher_name,'--', self.lessons_count)

    def teacher_graph(self):
        count_dates_list = []
        a = datetime.date.today()
        numdays = len(sorted(list(set(self.all_dates_array))))
        date_list = []
        for x in range(0, numdays): date_list.append(a - datetime.timedelta(days=x))
        date_list.reverse()
        count_dates_list.append(self.all_dates_array.count(date_list[0]))
        for date in range(1, len(date_list)):
            count_dates_list.append(self.all_dates_array.count(date_list[date]) + count_dates_list[date - 1])

        plt.plot(date_list, count_dates_list, label=self.teacher_name)
        plt.xlabel('DATES')
        plt.ylabel('INFO')


if __name__ == "__main__":
    test_key = {'Катерина Назаренко': '1nqqgT0WAbORvq6oHC2p4-afN87v8ZL-gDxQQmDkNr_o'}

    for key in KEYS_LIST:
        print('\n----------------------', key,
              '---------------------------')
        spreadsheet_data = get_spreadsheet_data.get_data(KEYS_LIST[key])
        teachers_data = Excel(spreadsheet_data, key)
        teachers_data.get_lesson_sum()
        # teachers_data.teacher_graph()

    # plt.title("Teachers all lessons")
    # plt.legend()
    # plt.show()
