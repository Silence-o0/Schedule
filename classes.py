from enum import IntEnum
import random
import pandas as pd


class Day(IntEnum):
    MONDAY = 1
    TUESDAY = 2
    WEDNESDAY = 3
    THURSDAY = 4
    FRIDAY = 5


class Group:
    def __init__(self, name, students_count, subjects):
        self.name = name
        self.students_count = students_count
        self.subjects = subjects


class SubjectDetail:
    def __init__(self, subj_type, hours=None, subgroups=None):
        self.type = subj_type
        self.hours = hours
        self.subgroups = subgroups


class Subject:
    def __init__(self, name, details):
        self.name = name
        self.details = details


class Teacher:
    def __init__(self, name, subjects, hours):
        self.name = name
        self.subjects = subjects
        self.hours = hours


class Auditorium:
    def __init__(self, number, capacity):
        self.number = number
        self.capacity = capacity


class Lesson:
    def __init__(
            self,
            day,
            lesson_num,
            teacher,
            lesson_type,
            subject,
            group,
            auditorium,
            subgroup=None,
    ):
        self.day = day
        self.lesson_num = lesson_num
        self.teacher = teacher.name
        self.lesson_type = lesson_type
        self.subject = subject.name
        self.group = group.name
        self.auditorium = auditorium.number
        self.subgroup = subgroup

    def to_dict(self):
        return {
            "Day": self.day.name,
            "Lesson_num": self.lesson_num,
            "Teacher": self.teacher,
            "Lesson_type": self.lesson_type,
            "Subject": self.subject,
            "Group": self.group,
            "Auditorium": self.auditorium,
            "Subgroup": self.subgroup,
        }


class Schedule:
    def __init__(self):
        self.lessons = []

    def _check_constraints(self, lesson):
        for existing_lesson in self.lessons:
            if (
                    existing_lesson.teacher == lesson.teacher
                    and existing_lesson.day == lesson.day
                    and existing_lesson.lesson_num == lesson.lesson_num
            ):
                return False
            if (
                    existing_lesson.group == lesson.group
                    and existing_lesson.day == lesson.day
                    and existing_lesson.lesson_num == lesson.lesson_num
                    and (existing_lesson.subgroup is None or lesson.subgroup is None or
                         existing_lesson.subgroup[:-1] != lesson.subgroup[:-1] or
                         existing_lesson.subgroup == lesson.subgroup)
            ):
                return False
            if (
                    existing_lesson.auditorium == lesson.auditorium
                    and existing_lesson.day == lesson.day
                    and existing_lesson.lesson_num == lesson.lesson_num
            ):
                if lesson.lesson_type == "Lec" or existing_lesson.lesson_type == "Lec":
                    return False
        return True

    def check_change_shared_lec(self, lesson):
        for existing_lesson in self.lessons:

                self.check_change_shared_lec(lesson)
                return True

    def check_shared_lec(self, group, subject):
        for existing_lesson in self.lessons:
            if existing_lesson.group != group and existing_lesson.subject == subject:
                lesson = Lesson(
                    existing_lesson.day,
                    existing_lesson.lesson_num,
                    existing_lesson.teacher,
                    existing_lesson.lesson_type,
                    subject,
                    group,
                    existing_lesson.auditorium,
                )
                self.check_change_shared_lec(lesson)
                return True
        return False

    def generate_schedule(self, groups, teachers, auditoriums, week_quantity):
        max_lessons_per_day = 4
        days = list(Day)

        for group in groups:
            for subject in group.subjects:
                for detail in subject.details:
                    hours = detail.hours
                    lesson_quantity = hours / 1.5
                    lesson_type = detail.type
                    subgroup_count = (
                        int(detail.subgroups) if lesson_type == "Lab" else 1
                    )
                    available_teachers = [
                        t
                        for t in teachers
                        if any(
                            ts.name == subject.name
                            and detail.type in [tsd.type for tsd in ts.details]
                            for ts in t.subjects
                        )
                    ]
                    if not available_teachers:
                        print(
                            f"No available teachers for {subject.name}, {lesson_type}, {group.name}"
                        )
                        break
                    teacher = random.choice(available_teachers)
                    for _ in range(int((lesson_quantity-1) // week_quantity) + 1):
                        # if lesson_type == "Lec":
                        #     if self.check_shared_lec(group, subject):
                        #         break
                        for subgroup in range(1, subgroup_count + 1):
                            subgroup = f"{subgroup}/{subgroup_count}"
                            assigned = False
                            loop_num = 0
                            while not assigned and loop_num < 100:
                                if subgroup_count > 1:
                                    teacher = random.choice(available_teachers)
                                day = random.choice(days)
                                lesson_num = random.randint(1, max_lessons_per_day)
                                auditorium = random.choice(
                                    [
                                        a
                                        for a in auditoriums
                                        if a.capacity >= group.students_count
                                    ]
                                )
                                if not auditorium:
                                    auditorium = random.choice(
                                        [
                                            a
                                            for a in auditoriums
                                        ]
                                    )

                                if lesson_type == "Lec" or subgroup_count == 1:
                                    subgroup = None

                                lesson = Lesson(
                                    day,
                                    lesson_num,
                                    teacher,
                                    lesson_type,
                                    subject,
                                    group,
                                    auditorium,
                                    subgroup,
                                )
                                loop_num += 1
                                if self._check_constraints(lesson):
                                    self.lessons.append(lesson)
                                    assigned = True


    def export_schedule_to_excel(self, filename="schedule.xlsx"):
        lesson_dicts = [
            {
                **lesson.to_dict(),
                "day_num": lesson.day.value,
            }
            for lesson in self.lessons
        ]
        df = pd.DataFrame(lesson_dicts)

        column_for_group_order = [
            "Group",
            "Day",
            "Lesson_num",
            "Subject",
            "Lesson_type",
            "Teacher",
            "Subgroup",
            "Auditorium",
        ]
        column_for_teacher_order = [
            "Teacher",
            "Day",
            "Lesson_num",
            "Subject",
            "Lesson_type",
            "Group",
            "Subgroup",
            "Auditorium",
        ]

        sorted_by_group = (
            df[column_for_group_order + ["day_num"]]
            .sort_values(by=["Group", "day_num", "Lesson_num"])
            .drop(columns=["day_num"])
        )

        sorted_by_teacher = (
            df[column_for_teacher_order + ["day_num"]]
            .sort_values(by=["Teacher", "day_num", "Lesson_num"])
            .drop(columns=["day_num"])
        )

        with pd.ExcelWriter(filename) as writer:
            sorted_by_group.to_excel(writer, sheet_name="Sorted_By_Groups", index=False)
            sorted_by_teacher.to_excel(
                writer, sheet_name="Sorted_By_Teachers", index=False
            )
        print(f"Saved in '{filename}'")