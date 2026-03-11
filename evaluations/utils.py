from .models import ActionLog

def log_action(user, action, object_type='', object_id=None, description=''):
    ActionLog.objects.create(
        user = user if user.is_authenticated else None,
        action = action,
        object_type = object_type,
        object_id = object_id,
        description = description
    )
