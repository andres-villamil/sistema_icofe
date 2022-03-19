class ddiRouter(object):
    """
    A router to control all database operations on models in the
    auth application.
    """
    def db_for_read(self, model, **hints):
        """
        Attempts to read demandas models go to  db sen_ddi.
        """
        if model._meta.app_label == 'demandas':
            return 'sen_ddi'
        return None

    def db_for_write(self, model, **hints):
        """
        Attempts to write demandas models go to db sen_ddi.
        """
        if model._meta.app_label == 'demandas':
            return 'sen_ddi'
        return None

    def allow_relation(self, obj1, obj2, **hints):
        """
        Allow relations if a model in the demandas app is involved.
        """
        if obj1._meta.app_label == 'demandas' or \
           obj2._meta.app_label == 'demandas':
           return True
        return None

    def allow_migrate(self, db, app_label, model_name=None, **hints):
        """
        Make sure the demandas app only appears in the db 'sen_ddi'
        database.
        """
        if app_label == 'demandas':
            return db == 'sen_ddi'
        return None